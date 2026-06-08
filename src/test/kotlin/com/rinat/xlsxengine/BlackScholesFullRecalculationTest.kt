package com.rinat.xlsxengine

import java.nio.file.Path
import kotlin.io.path.Path
import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertTrue

class BlackScholesFullRecalculationTest {
    private val workbookFile = "BlackScholes.xlsx"

    @Test
    fun `should recalculate all cells on all sheets for black scholes workbook`() {
        val path = resourcePath(workbookFile)
        val workbook = XlsxReader.read(path)
        val engine = XlsxFormulaEngine.fromFile(path)
        val evaluated = engine.evaluateAll()

        var totalCells = 0
        var matchedCells = 0
        var formulaCells = 0
        var matchedFormulaCells = 0
        val mismatches = mutableListOf<String>()

        for (sheetName in workbook.sheetOrder) {
            val sheet = workbook.sheets.getValue(sheetName)
            val evaluatedSheet = evaluated.sheets.getValue(sheetName)

            for (cell in sheet.cells.values.sortedWith(compareBy<WorkbookCell>({ it.address.row }, { it.address.column }))) {
                totalCells += 1
                if (!cell.formula.isNullOrBlank()) formulaCells += 1

                val expected = expectedValue(workbook, cell)
                val actual = evaluatedSheet.values[cell.address]
                    ?: engine.evaluateCell(sheetName, cell.address.toA1())

                val ok = isEquivalent(expected, actual)
                if (ok) {
                    matchedCells += 1
                    if (!cell.formula.isNullOrBlank()) matchedFormulaCells += 1
                } else {
                    mismatches += "$sheetName!${cell.address.toA1()} formula=${cell.formula ?: "<literal>"} expected=$expected actual=$actual"
                }
            }
        }

        val overallAccuracy = if (totalCells == 0) 1.0 else matchedCells.toDouble() / totalCells
        val formulaAccuracy = if (formulaCells == 0) 1.0 else matchedFormulaCells.toDouble() / formulaCells

        assertTrue(totalCells > 0, "Workbook must contain cells")
        assertTrue(
            mismatches.isEmpty(),
            "BlackScholes recalculation mismatches=${mismatches.size}, totalCells=$totalCells, " +
                "overallAccuracy=$overallAccuracy, formulaCells=$formulaCells, formulaAccuracy=$formulaAccuracy. " +
                "First mismatches: ${mismatches.take(20)}"
        )
    }

    private fun expectedValue(workbook: WorkbookData, cell: WorkbookCell): CellValue {
        val raw = cell.rawValue ?: return CellValue.Blank
        return when (cell.type) {
            "s" -> {
                val idx = raw.toIntOrNull() ?: return CellValue.ErrorValue("Invalid shared string index '$raw'")
                workbook.sharedStrings.getOrNull(idx)?.let { CellValue.TextValue(it) }
                    ?: CellValue.ErrorValue("Shared string index out of bounds: $idx")
            }

            "str", "inlineStr" -> CellValue.TextValue(raw)
            "b" -> CellValue.BooleanValue(raw == "1")
            "e" -> CellValue.ErrorValue(raw)
            else -> raw.toDoubleOrNull()?.let { CellValue.NumberValue(it) } ?: CellValue.TextValue(raw)
        }
    }

    private fun isEquivalent(expected: CellValue, actual: CellValue): Boolean {
        return when {
            expected is CellValue.NumberValue && actual is CellValue.NumberValue -> {
                abs(expected.value - actual.value) <= 1e-4
            }

            expected is CellValue.BooleanValue && actual is CellValue.BooleanValue -> expected.value == actual.value
            expected is CellValue.TextValue && actual is CellValue.TextValue -> expected.value == actual.value
            expected is CellValue.ErrorValue && actual is CellValue.ErrorValue -> true
            expected is CellValue.Blank && actual is CellValue.Blank -> true
            else -> false
        }
    }

    private fun resourcePath(fileName: String): Path = Path("src/test/resources/$fileName")
}
