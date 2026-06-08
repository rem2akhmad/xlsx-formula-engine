package com.rinat.xlsxengine

import kotlin.math.pow

internal interface FormulaEvaluator {
    fun evaluateCell(sheet: String, address: CellAddress): CellValue
    fun evaluateSheet(sheet: String): EvaluatedSheet
}

internal class DefaultFormulaEvaluator(
    private val workbook: WorkbookData,
    private val functionResolver: (String) -> FormulaFunctionHandler? = FunctionRegistry::resolve
) : FormulaEvaluator, FunctionEvaluationContext {
    private data class CellKey(val sheet: String, val address: CellAddress)

    private val cache = mutableMapOf<CellKey, CellValue>()
    private val visiting = mutableSetOf<CellKey>()
    private val resolvingNames = mutableSetOf<Pair<String, String>>()

    override fun evaluateCell(sheet: String, address: CellAddress): CellValue {
        val key = CellKey(sheet, address)
        cache[key]?.let { return it }

        if (!visiting.add(key)) {
            return CellValue.ErrorValue(ErrorCodes.REF)
        }

        val value = try {
            val sheetData = workbook.sheets[sheet]
            if (sheetData == null) {
                CellValue.ErrorValue(ErrorCodes.REF)
            } else {
                val cell = sheetData.cells[address]
                if (cell == null) {
                    CellValue.Blank
                } else if (!cell.formula.isNullOrBlank()) {
                    val node = parseFormula(cell.formula)
                    val formulaValue = when (val eval = evalNode(sheet, node)) {
                        is EvalValue.Scalar -> eval.value
                        is EvalValue.Range, is EvalValue.Range3D -> CellValue.ErrorValue(ErrorCodes.VALUE)
                    }
                    if (formulaValue == CellValue.Blank) CellValue.NumberValue(0.0) else formulaValue
                } else {
                    parseLiteralCellValue(cell)
                }
            }
        } catch (_: Throwable) {
            CellValue.ErrorValue(ErrorCodes.VALUE)
        }

        visiting.remove(key)
        cache[key] = value
        return value
    }

    override fun evaluateSheet(sheet: String): EvaluatedSheet {
        val sheetData = workbook.sheets[sheet] ?: error("Unknown sheet '$sheet'")
        val out = LinkedHashMap<CellAddress, CellValue>()
        for (address in sheetData.cells.keys.sortedWith(compareBy<CellAddress>({ it.row }, { it.column }))) {
            out[address] = evaluateCell(sheet, address)
        }
        return EvaluatedSheet(name = sheet, values = out)
    }

    override fun evalNode(currentSheet: String, node: FormulaNode): EvalValue {
        return when (node) {
            is FormulaNode.NumberLiteral -> EvalValue.Scalar(CellValue.NumberValue(node.value))
            is FormulaNode.StringLiteral -> EvalValue.Scalar(CellValue.TextValue(node.value))
            is FormulaNode.BooleanLiteral -> EvalValue.Scalar(CellValue.BooleanValue(node.value))
            is FormulaNode.ErrorLiteral -> EvalValue.Scalar(CellValue.ErrorValue(node.code))
            FormulaNode.BlankLiteral -> EvalValue.Scalar(CellValue.Blank)

            is FormulaNode.UnaryOp -> {
                val number = evalScalar(currentSheet, node.expr).asNumberOrNull()
                    ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                when (node.op) {
                    "+" -> EvalValue.Scalar(CellValue.NumberValue(number))
                    "-" -> EvalValue.Scalar(CellValue.NumberValue(-number))
                    else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                }
            }

            is FormulaNode.PostfixOp -> {
                when (node.op) {
                    "%" -> {
                        val number = evalScalar(currentSheet, node.expr).asNumberOrNull()
                            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                        EvalValue.Scalar(CellValue.NumberValue(number / 100.0))
                    }

                    else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                }
            }

            is FormulaNode.BinaryOp -> evalBinary(currentSheet, node)

            is FormulaNode.CellRef -> {
                val sheet = node.sheet ?: currentSheet
                EvalValue.Scalar(evaluateCell(sheet, node.ref))
            }

            is FormulaNode.RangeRef -> {
                val sheet = node.sheet ?: currentSheet
                val normalized = normalizeRange(node.start, node.end)
                EvalValue.Range(sheet, normalized.first, normalized.second)
            }

            is FormulaNode.Range3DRef -> {
                val sheetSpan = sheetSpan(node.startSheet, node.endSheet)
                    ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
                val normalized = normalizeRange(node.start, node.end)
                EvalValue.Range3D(
                    ranges = sheetSpan.map { sheetName ->
                        EvalValue.Range(sheetName, normalized.first, normalized.second)
                    }
                )
            }

            is FormulaNode.DefinedNameRef -> evalDefinedName(currentSheet, node.name)
            is FormulaNode.FunctionCall -> evalFunction(currentSheet, node)
        }
    }

    override fun evalScalar(currentSheet: String, node: FormulaNode): CellValue {
        return when (val value = evalNode(currentSheet, node)) {
            is EvalValue.Scalar -> value.value
            is EvalValue.Range, is EvalValue.Range3D -> CellValue.ErrorValue(ErrorCodes.VALUE)
        }
    }

    override fun iterateRange(sheet: String, start: CellAddress, end: CellAddress): Sequence<CellValue> {
        return sequence {
            for (row in start.row..end.row) {
                for (col in start.column..end.column) {
                    yield(evaluateCell(sheet, CellAddress(row, col)))
                }
            }
        }
    }

    override fun forEachArgValue(
        currentSheet: String,
        args: List<FormulaNode>,
        onValue: (CellValue) -> Unit
    ): CellValue.ErrorValue? {
        for (arg in args) {
            when (val value = evalNode(currentSheet, arg)) {
                is EvalValue.Scalar -> {
                    if (value.value is CellValue.ErrorValue) return value.value
                    onValue(value.value)
                }

                is EvalValue.Range -> {
                    for (cellValue in iterateRange(value.sheet, value.start, value.end)) {
                        if (cellValue is CellValue.ErrorValue) return cellValue
                        onValue(cellValue)
                    }
                }

                is EvalValue.Range3D -> {
                    for (range in value.ranges) {
                        for (cellValue in iterateRange(range.sheet, range.start, range.end)) {
                            if (cellValue is CellValue.ErrorValue) return cellValue
                            onValue(cellValue)
                        }
                    }
                }
            }
        }
        return null
    }

    override fun collectArgValues(currentSheet: String, args: List<FormulaNode>): List<CellValue> {
        val values = mutableListOf<CellValue>()
        forEachArgValue(currentSheet, args) { value ->
            values += value
        }
        return values
    }

    override fun asBoolean(value: CellValue): Boolean? {
        return when (value) {
            is CellValue.BooleanValue -> value.value
            is CellValue.NumberValue -> value.value != 0.0
            CellValue.Blank -> false
            is CellValue.TextValue -> {
                when {
                    value.value.equals("TRUE", ignoreCase = true) -> true
                    value.value.equals("FALSE", ignoreCase = true) -> false
                    value.value.toDoubleOrNull() != null -> value.value.toDouble() != 0.0
                    else -> null
                }
            }

            is CellValue.ErrorValue -> null
        }
    }

    private fun parseLiteralCellValue(cell: WorkbookCell): CellValue {
        val raw = cell.rawValue ?: return CellValue.Blank
        return when (cell.type) {
            "s" -> {
                val idx = raw.toIntOrNull() ?: return CellValue.ErrorValue(ErrorCodes.VALUE)
                val text = workbook.sharedStrings.getOrNull(idx) ?: return CellValue.ErrorValue(ErrorCodes.REF)
                CellValue.TextValue(text)
            }

            "str", "inlineStr" -> CellValue.TextValue(raw)
            "b" -> CellValue.BooleanValue(raw == "1")
            "e" -> CellValue.ErrorValue(raw)
            else -> raw.toDoubleOrNull()?.let(CellValue::NumberValue) ?: CellValue.TextValue(raw)
        }
    }

    private fun evalBinary(currentSheet: String, node: FormulaNode.BinaryOp): EvalValue {
        if (node.op == "&") {
            val left = evalScalar(currentSheet, node.left)
            val right = evalScalar(currentSheet, node.right)
            if (left is CellValue.ErrorValue) return EvalValue.Scalar(left)
            if (right is CellValue.ErrorValue) return EvalValue.Scalar(right)
            return EvalValue.Scalar(CellValue.TextValue(toText(left) + toText(right)))
        }

        val left = evalScalar(currentSheet, node.left)
        val right = evalScalar(currentSheet, node.right)

        if (left is CellValue.ErrorValue) return EvalValue.Scalar(left)
        if (right is CellValue.ErrorValue) return EvalValue.Scalar(right)

        return when (node.op) {
            "+", "-", "*", "/", "^" -> {
                val l = left.asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                val r = right.asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

                val result = when (node.op) {
                    "+" -> l + r
                    "-" -> l - r
                    "*" -> l * r
                    "/" -> if (r == 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0)) else l / r
                    "^" -> l.pow(r)
                    else -> error("unreachable")
                }
                EvalValue.Scalar(CellValue.NumberValue(result))
            }

            ">", "<", ">=", "<=", "=", "<>" -> {
                val result = compareValues(left, right, node.op)
                EvalValue.Scalar(CellValue.BooleanValue(result))
            }

            else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }

    private fun compareValues(left: CellValue, right: CellValue, op: String): Boolean {
        val lNumber = left.asNumberOrNull()
        val rNumber = right.asNumberOrNull()
        if (lNumber != null && rNumber != null) {
            return when (op) {
                ">" -> lNumber > rNumber
                "<" -> lNumber < rNumber
                ">=" -> lNumber >= rNumber
                "<=" -> lNumber <= rNumber
                "=" -> lNumber == rNumber
                "<>" -> lNumber != rNumber
                else -> false
            }
        }

        val lText = toText(left)
        val rText = toText(right)
        return when (op) {
            "=" -> lText == rText
            "<>" -> lText != rText
            ">" -> lText > rText
            "<" -> lText < rText
            ">=" -> lText >= rText
            "<=" -> lText <= rText
            else -> false
        }
    }

    private fun toText(value: CellValue): String {
        return when (value) {
            is CellValue.TextValue -> value.value
            is CellValue.BooleanValue -> if (value.value) "TRUE" else "FALSE"
            is CellValue.NumberValue -> value.value.toString()
            CellValue.Blank -> ""
            is CellValue.ErrorValue -> value.message
        }
    }

    private fun evalFunction(currentSheet: String, call: FormulaNode.FunctionCall): EvalValue {
        val handler = functionResolver(call.name)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NAME))
        return handler.evaluate(this, currentSheet, call.args)
    }

    private fun evalDefinedName(currentSheet: String, name: String): EvalValue {
        val normalizedName = name.uppercase()
        val key = currentSheet to normalizedName
        if (!resolvingNames.add(key)) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        }

        return try {
            val formula = workbook.localDefinedNames[currentSheet]?.get(normalizedName)
                ?: workbook.definedNames[normalizedName]
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NAME))
            evalNode(currentSheet, parseFormula(formula))
        } catch (_: Throwable) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } finally {
            resolvingNames.remove(key)
        }
    }

    private fun sheetSpan(startSheet: String, endSheet: String): List<String>? {
        val startIdx = workbook.sheetOrder.indexOf(startSheet)
        val endIdx = workbook.sheetOrder.indexOf(endSheet)
        if (startIdx < 0 || endIdx < 0) return null
        return if (startIdx <= endIdx) {
            workbook.sheetOrder.subList(startIdx, endIdx + 1)
        } else {
            workbook.sheetOrder.subList(endIdx, startIdx + 1).reversed()
        }
    }

    private fun normalizeRange(a: CellAddress, b: CellAddress): Pair<CellAddress, CellAddress> {
        val start = CellAddress(minOf(a.row, b.row), minOf(a.column, b.column))
        val end = CellAddress(maxOf(a.row, b.row), maxOf(a.column, b.column))
        return start to end
    }
}
