package com.rinat.xlsxengine

import java.nio.file.Path

class XlsxFormulaEngine private constructor(
    private val workbook: WorkbookData
) {
    private var currentWorkbook: WorkbookData = workbook
    private var evaluator: FormulaEvaluator = DefaultFormulaEvaluator(currentWorkbook)

    fun evaluateAll(): EvaluatedWorkbook {
        // Pre-warm according to calcChain when present to match workbook-defined recalc order.
        for (entry in currentWorkbook.calcChain) {
            evaluator.evaluateCell(entry.sheet, entry.address)
        }
        val out = LinkedHashMap<String, EvaluatedSheet>()
        for (sheet in currentWorkbook.sheetOrder) {
            out[sheet] = evaluator.evaluateSheet(sheet)
        }
        return EvaluatedWorkbook(out)
    }

    fun evaluateCell(sheet: String, addressA1: String): CellValue {
        return evaluator.evaluateCell(sheet, CellAddress.parse(addressA1))
    }

    fun evaluateRow(sheet: String, rowNumber: Int): Map<String, CellValue> {
        require(rowNumber >= 1) { "Row number must be >= 1" }
        val sheetData = currentWorkbook.sheets[sheet] ?: error("Unknown sheet '$sheet'")
        return sheetData.cells.keys
            .asSequence()
            .filter { it.row == rowNumber }
            .sortedBy { it.column }
            .associate { address -> address.toA1() to evaluator.evaluateCell(sheet, address) }
    }

    fun evaluateColumn(sheet: String, columnA1Name: String): Map<Int, CellValue> {
        val column = CellAddress.nameToColumn(columnA1Name)
        val sheetData = currentWorkbook.sheets[sheet] ?: error("Unknown sheet '$sheet'")
        return sheetData.cells.keys
            .asSequence()
            .filter { it.column == column }
            .sortedBy { it.row }
            .associate { address -> address.row to evaluator.evaluateCell(sheet, address) }
    }

    fun setCellNumber(sheet: String, addressA1: String, value: Double) {
        updateLiteralCell(sheet, addressA1, type = null, rawValue = value.toString())
    }

    fun setCellText(sheet: String, addressA1: String, value: String) {
        updateLiteralCell(sheet, addressA1, type = "str", rawValue = value)
    }

    fun setCellBoolean(sheet: String, addressA1: String, value: Boolean) {
        updateLiteralCell(sheet, addressA1, type = "b", rawValue = if (value) "1" else "0")
    }

    fun clearCell(sheet: String, addressA1: String) {
        val address = CellAddress.parse(addressA1)
        val sheetData = currentWorkbook.sheets[sheet] ?: error("Unknown sheet '$sheet'")
        val updatedCells = sheetData.cells.toMutableMap()
        updatedCells.remove(address)
        replaceSheet(sheetData.copy(cells = updatedCells))
    }

    private fun updateLiteralCell(sheet: String, addressA1: String, type: String?, rawValue: String?) {
        val address = CellAddress.parse(addressA1)
        val sheetData = currentWorkbook.sheets[sheet] ?: error("Unknown sheet '$sheet'")
        val updatedCells = sheetData.cells.toMutableMap()
        updatedCells[address] = WorkbookCell(
            address = address,
            type = type,
            rawValue = rawValue,
            formula = null
        )
        replaceSheet(sheetData.copy(cells = updatedCells))
    }

    private fun replaceSheet(updatedSheet: WorksheetData) {
        val updatedSheets = currentWorkbook.sheets.toMutableMap()
        updatedSheets[updatedSheet.name] = updatedSheet
        currentWorkbook = currentWorkbook.copy(sheets = updatedSheets)
        evaluator = DefaultFormulaEvaluator(currentWorkbook)
    }

    companion object {
        fun fromFile(path: Path): XlsxFormulaEngine = XlsxFormulaEngine(XlsxReader.read(path))
    }
}
