package com.rinat.xlsxengine

import java.io.InputStream
import java.net.URI
import java.nio.file.Path
import java.util.zip.ZipFile
import javax.xml.stream.XMLInputFactory
import javax.xml.stream.XMLStreamConstants

object XlsxReader {
    private val xmlFactory: XMLInputFactory = XMLInputFactory.newFactory().apply {
        runCatching { setProperty(XMLInputFactory.IS_NAMESPACE_AWARE, true) }
        runCatching { setProperty(XMLInputFactory.SUPPORT_DTD, false) }
        runCatching { setProperty("javax.xml.stream.isSupportingExternalEntities", false) }
    }

    fun read(path: Path): WorkbookData {
        ZipFile(path.toFile()).use { zip ->
            val workbookEntry = zip.getEntry("xl/workbook.xml")
                ?: error("Invalid xlsx: missing xl/workbook.xml")
            val workbookRelsEntry = zip.getEntry("xl/_rels/workbook.xml.rels")
                ?: error("Invalid xlsx: missing xl/_rels/workbook.xml.rels")

            val workbookMeta = zip.getInputStream(workbookEntry).use(::parseWorkbook)
            val relationships = zip.getInputStream(workbookRelsEntry).use(::parseRelationships)
            val sharedStrings = zip.getEntry("xl/sharedStrings.xml")
                ?.let { entry -> zip.getInputStream(entry).use(::parseSharedStrings) }
                ?: emptyList()
            val calcChain = zip.getEntry("xl/calcChain.xml")
                ?.let { entry ->
                    zip.getInputStream(entry).use { input ->
                        parseCalcChain(input, workbookMeta.sheets)
                    }
                }
                ?: emptyList()

            val sheetMap = LinkedHashMap<String, WorksheetData>()
            for (sheet in workbookMeta.sheets) {
                val relationship = relationships[sheet.relationshipId]
                    ?: error("Missing relationship target for ${sheet.relationshipId}")
                require(!relationship.external) { "External worksheet relationship is not supported: ${relationship.target}" }
                val normalizedPath = resolveRelationshipTarget("xl/workbook.xml", relationship.target)
                val entry = zip.getEntry(normalizedPath)
                    ?: error("Missing worksheet file: $normalizedPath")
                val sheetData = zip.getInputStream(entry).use { stream ->
                    parseWorksheet(sheet.name, stream)
                }
                sheetMap[sheet.name] = sheetData
            }

            return WorkbookData(
                sheets = sheetMap,
                sheetOrder = workbookMeta.sheets.map { it.name },
                sharedStrings = sharedStrings,
                calcChain = calcChain,
                definedNames = workbookMeta.definedNames,
                localDefinedNames = workbookMeta.localDefinedNames
            )
        }
    }

    private fun resolveRelationshipTarget(sourcePart: String, target: String): String {
        if (target.startsWith("/")) return target.removePrefix("/")
        val sourceDir = sourcePart.substringBeforeLast('/', missingDelimiterValue = "") + "/"
        val base = URI("https://xlsx.local/$sourceDir")
        val resolved = base.resolve(target).normalize()
        return resolved.path.removePrefix("/")
    }

    private data class SheetDef(val name: String, val relationshipId: String, val sheetId: Int?)
    private data class WorkbookMeta(
        val sheets: List<SheetDef>,
        val definedNames: Map<String, String>,
        val localDefinedNames: Map<String, Map<String, String>>
    )
    private data class RelationshipTarget(val target: String, val external: Boolean)

    private data class SharedFormulaTemplate(
        val baseAddress: CellAddress,
        val baseFormula: String
    )

    private fun parseWorkbook(input: InputStream): WorkbookMeta {
        val reader = xmlFactory.createXMLStreamReader(input)
        try {
            val sheets = mutableListOf<SheetDef>()
            val definedNames = linkedMapOf<String, String>()
            val localDefinedNames = mutableMapOf<Int, MutableMap<String, String>>()

            var inDefinedName = false
            var currentDefinedName: String? = null
            var currentLocalSheetId: Int? = null
            var definedNameText: StringBuilder? = null

            while (reader.hasNext()) {
                when (reader.next()) {
                    XMLStreamConstants.START_ELEMENT -> {
                        when (reader.localName) {
                            "sheet" -> {
                                val name = reader.getAttributeValue(null, "name")
                                    ?: error("sheet name missing")
                                val relId = reader.getAttributeValue(
                                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                                    "id"
                                ) ?: reader.getAttributeValue(null, "id")
                                ?: error("sheet relationship id missing")
                                val sheetId = reader.getAttributeValue(null, "sheetId")?.toIntOrNull()
                                sheets += SheetDef(name, relId, sheetId)
                            }

                            "definedName" -> {
                                inDefinedName = true
                                currentDefinedName = reader.getAttributeValue(null, "name")
                                currentLocalSheetId = reader.getAttributeValue(null, "localSheetId")?.toIntOrNull()
                                definedNameText = StringBuilder()
                            }
                        }
                    }

                    XMLStreamConstants.CHARACTERS, XMLStreamConstants.CDATA -> {
                        if (inDefinedName) {
                            definedNameText?.append(reader.text)
                        }
                    }

                    XMLStreamConstants.END_ELEMENT -> {
                        if (reader.localName == "definedName") {
                            val name = currentDefinedName?.trim()
                            val value = definedNameText?.toString()?.trim().orEmpty()
                            if (!name.isNullOrBlank() && value.isNotBlank()) {
                                val localSheetId = currentLocalSheetId
                                if (localSheetId == null) {
                                    definedNames[name.uppercase()] = value
                                } else {
                                    val bucket = localDefinedNames.getOrPut(localSheetId) { linkedMapOf() }
                                    bucket[name.uppercase()] = value
                                }
                            }
                            inDefinedName = false
                            currentDefinedName = null
                            currentLocalSheetId = null
                            definedNameText = null
                        }
                    }
                }
            }

            val localBySheetName = mutableMapOf<String, Map<String, String>>()
            for ((localSheetId, names) in localDefinedNames) {
                val sheet = sheets.getOrNull(localSheetId) ?: continue
                localBySheetName[sheet.name] = names
            }

            return WorkbookMeta(
                sheets = sheets,
                definedNames = definedNames,
                localDefinedNames = localBySheetName
            )
        } finally {
            reader.close()
        }
    }

    private fun parseRelationships(input: InputStream): Map<String, RelationshipTarget> {
        val reader = xmlFactory.createXMLStreamReader(input)
        try {
            val out = mutableMapOf<String, RelationshipTarget>()
            while (reader.hasNext()) {
                when (reader.next()) {
                    XMLStreamConstants.START_ELEMENT -> {
                        if (reader.localName == "Relationship") {
                            val id = reader.getAttributeValue(null, "Id") ?: continue
                            val target = reader.getAttributeValue(null, "Target") ?: continue
                            val targetMode = reader.getAttributeValue(null, "TargetMode")
                            out[id] = RelationshipTarget(
                                target = target,
                                external = targetMode.equals("External", ignoreCase = true)
                            )
                        }
                    }
                }
            }
            return out
        } finally {
            reader.close()
        }
    }

    private fun parseCalcChain(input: InputStream, sheets: List<SheetDef>): List<CalcChainEntry> {
        val reader = xmlFactory.createXMLStreamReader(input)
        try {
            val out = mutableListOf<CalcChainEntry>()
            var currentSheetName = sheets.firstOrNull()?.name
            while (reader.hasNext()) {
                when (reader.next()) {
                    XMLStreamConstants.START_ELEMENT -> {
                        if (reader.localName == "c") {
                            val sheetIndex = reader.getAttributeValue(null, "i")?.toIntOrNull()
                            if (sheetIndex != null) {
                                currentSheetName = sheets.getOrNull(sheetIndex - 1)?.name ?: currentSheetName
                            }
                            val ref = reader.getAttributeValue(null, "r")
                            if (!ref.isNullOrBlank() && currentSheetName != null) {
                                out += CalcChainEntry(
                                    sheet = currentSheetName,
                                    address = CellAddress.parse(ref)
                                )
                            }
                        }
                    }
                }
            }
            return out
        } finally {
            reader.close()
        }
    }

    private fun parseSharedStrings(input: InputStream): List<String> {
        val reader = xmlFactory.createXMLStreamReader(input)
        try {
            val out = mutableListOf<String>()
            var current = StringBuilder()
            var insideSi = false
            var insideT = false

            while (reader.hasNext()) {
                when (reader.next()) {
                    XMLStreamConstants.START_ELEMENT -> {
                        when (reader.localName) {
                            "si" -> {
                                insideSi = true
                                current = StringBuilder()
                            }

                            "t" -> if (insideSi) insideT = true
                        }
                    }

                    XMLStreamConstants.CHARACTERS, XMLStreamConstants.CDATA -> {
                        if (insideSi && insideT) {
                            current.append(reader.text)
                        }
                    }

                    XMLStreamConstants.END_ELEMENT -> {
                        when (reader.localName) {
                            "t" -> insideT = false
                            "si" -> {
                                insideSi = false
                                out += current.toString()
                            }
                        }
                    }
                }
            }
            return out
        } finally {
            reader.close()
        }
    }

    private fun parseWorksheet(name: String, input: InputStream): WorksheetData {
        val reader = xmlFactory.createXMLStreamReader(input)
        try {
            val cells = LinkedHashMap<CellAddress, WorkbookCell>()
            val sharedFormulaTemplates = mutableMapOf<Int, SharedFormulaTemplate>()

            var currentRef: String? = null
            var currentType: String? = null
            var currentFormula: String? = null
            var currentFormulaType: String? = null
            var currentFormulaSharedIndex: Int? = null
            var currentValue: String? = null
            var currentInlineText: StringBuilder? = null

            var textBuffer: StringBuilder? = null
            var currentTextTag: String? = null

            while (reader.hasNext()) {
                when (reader.next()) {
                    XMLStreamConstants.START_ELEMENT -> {
                        when (reader.localName) {
                            "c" -> {
                                currentRef = reader.getAttributeValue(null, "r")
                                currentType = reader.getAttributeValue(null, "t")
                                currentFormula = null
                                currentFormulaType = null
                                currentFormulaSharedIndex = null
                                currentValue = null
                                currentInlineText = if (currentType == "inlineStr") StringBuilder() else null
                            }

                            "f" -> {
                                currentFormulaType = reader.getAttributeValue(null, "t")
                                currentFormulaSharedIndex = reader.getAttributeValue(null, "si")?.toIntOrNull()
                                currentTextTag = "f"
                                textBuffer = StringBuilder()
                            }

                            "v" -> {
                                currentTextTag = "v"
                                textBuffer = StringBuilder()
                            }

                            "t" -> {
                                if (currentType == "inlineStr") {
                                    currentTextTag = "inlineT"
                                    textBuffer = StringBuilder()
                                }
                            }
                        }
                    }

                    XMLStreamConstants.CHARACTERS, XMLStreamConstants.CDATA -> {
                        textBuffer?.append(reader.text)
                    }

                    XMLStreamConstants.END_ELEMENT -> {
                        when (reader.localName) {
                            "f" -> {
                                if (currentTextTag == "f") {
                                    currentFormula = textBuffer?.toString()?.trim().orEmpty().ifBlank { null }
                                }
                                currentTextTag = null
                                textBuffer = null
                            }

                            "v" -> {
                                if (currentTextTag == "v") {
                                    currentValue = textBuffer?.toString()
                                }
                                currentTextTag = null
                                textBuffer = null
                            }

                            "t" -> {
                                if (currentTextTag == "inlineT") {
                                    currentInlineText?.append(textBuffer?.toString().orEmpty())
                                }
                                currentTextTag = null
                                textBuffer = null
                            }

                            "c" -> {
                                val ref = currentRef
                                if (!ref.isNullOrBlank()) {
                                    val address = CellAddress.parse(ref)
                                    val resolvedFormula = resolveSharedFormula(
                                        address = address,
                                        formula = currentFormula,
                                        formulaType = currentFormulaType,
                                        sharedIndex = currentFormulaSharedIndex,
                                        sharedFormulaTemplates = sharedFormulaTemplates
                                    )

                                    val rawValue = if (currentType == "inlineStr") {
                                        currentInlineText?.toString()
                                    } else {
                                        currentValue
                                    }

                                    cells[address] = WorkbookCell(
                                        address = address,
                                        type = currentType,
                                        rawValue = rawValue,
                                        formula = resolvedFormula
                                    )
                                }

                                currentRef = null
                                currentType = null
                                currentFormula = null
                                currentFormulaType = null
                                currentFormulaSharedIndex = null
                                currentValue = null
                                currentInlineText = null
                            }
                        }
                    }
                }
            }

            return WorksheetData(name = name, cells = cells)
        } finally {
            reader.close()
        }
    }

    private fun resolveSharedFormula(
        address: CellAddress,
        formula: String?,
        formulaType: String?,
        sharedIndex: Int?,
        sharedFormulaTemplates: MutableMap<Int, SharedFormulaTemplate>
    ): String? {
        if (formulaType != "shared") return formula
        val si = sharedIndex ?: return formula

        if (!formula.isNullOrBlank()) {
            val normalized = formula.trim()
            sharedFormulaTemplates[si] = SharedFormulaTemplate(address, normalized)
            return normalized
        }

        val template = sharedFormulaTemplates[si] ?: return null
        val rowDelta = address.row - template.baseAddress.row
        val colDelta = address.column - template.baseAddress.column
        return shiftFormulaReferences(template.baseFormula, rowDelta, colDelta)
    }

    private fun shiftFormulaReferences(formula: String, rowDelta: Int, colDelta: Int): String {
        if (rowDelta == 0 && colDelta == 0) return formula

        val out = StringBuilder()
        var i = 0
        while (i < formula.length) {
            val ch = formula[i]

            if (ch == '"') {
                val start = i
                i++
                while (i < formula.length) {
                    if (formula[i] == '"') {
                        if (i + 1 < formula.length && formula[i + 1] == '"') {
                            i += 2
                            continue
                        }
                        i++
                        break
                    }
                    i++
                }
                out.append(formula.substring(start, i))
                continue
            }

            if (ch == '\'') {
                val start = i
                i++
                while (i < formula.length) {
                    if (formula[i] == '\'') {
                        if (i + 1 < formula.length && formula[i + 1] == '\'') {
                            i += 2
                            continue
                        }
                        i++
                        break
                    }
                    i++
                }
                out.append(formula.substring(start, i))
                continue
            }

            val shifted = tryShiftCellReference(formula, i, rowDelta, colDelta)
            if (shifted != null) {
                out.append(shifted.first)
                i = shifted.second
                continue
            }

            out.append(ch)
            i++
        }

        return out.toString()
    }

    private fun tryShiftCellReference(
        source: String,
        startIndex: Int,
        rowDelta: Int,
        colDelta: Int
    ): Pair<String, Int>? {
        if (startIndex > 0) {
            val prev = source[startIndex - 1]
            if (prev.isLetterOrDigit() || prev == '_' || prev == '.') {
                return null
            }
        }

        var i = startIndex
        var colAbs = false
        var rowAbs = false

        if (i < source.length && source[i] == '$') {
            colAbs = true
            i++
        }

        val colStart = i
        while (i < source.length && source[i].isLetter()) i++
        val col = source.substring(colStart, i)
        if (col.length !in 1..3) return null

        if (i < source.length && source[i] == '$') {
            rowAbs = true
            i++
        }

        val rowStart = i
        while (i < source.length && source[i].isDigit()) i++
        if (rowStart == i) return null

        if (i < source.length) {
            val next = source[i]
            if (next.isLetterOrDigit() || next == '_' || next == '.') {
                return null
            }
        }

        val row = source.substring(rowStart, i).toIntOrNull() ?: return null
        var colIndex = CellAddress.nameToColumn(col)
        var rowIndex = row

        if (!colAbs) colIndex += colDelta
        if (!rowAbs) rowIndex += rowDelta

        if (colIndex < 1 || rowIndex < 1) {
            return "#REF!" to i
        }

        val shifted = buildString {
            if (colAbs) append('$')
            append(CellAddress.columnToName(colIndex))
            if (rowAbs) append('$')
            append(rowIndex)
        }

        return shifted to i
    }
}
