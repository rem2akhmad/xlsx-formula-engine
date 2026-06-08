package com.rinat.xlsxengine

import kotlin.io.path.Path
import kotlin.math.abs
import kotlin.math.max
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class ComplexFormulaMutationTest {
    @Test
    fun `savings modeling sheet3 f3 should follow IF branch logic after C3 mutation`() {
        val path = Path("src/test/resources/Savings modeling.xlsx")
        val engine = XlsxFormulaEngine.fromFile(path)

        val d3 = numeric(engine.evaluateCell("Sheet3", "D3"))
        val e3 = numeric(engine.evaluateCell("Sheet3", "E3"))

        engine.setCellNumber("Sheet3", "C3", d3 - 1.0)
        val whenTrue = numeric(engine.evaluateCell("Sheet3", "F3"))
        val expectedTrue = e3 * d3 + 201.0
        assertClose(expectedTrue, whenTrue, "Sheet3!F3 true branch")

        engine.setCellNumber("Sheet3", "C3", d3 + 1.0)
        val whenFalse = numeric(engine.evaluateCell("Sheet3", "F3"))
        val expectedFalse = e3 * d3 - 1.0
        assertClose(expectedFalse, whenFalse, "Sheet3!F3 false branch")
    }

    @Test
    fun `cfi workbook d3 should follow IFERROR-IF-ABS logic after D60 mutation`() {
        val path = Path("src/test/resources/CFI-Case-Study-Three-Statement-Model.xlsx")
        val engine = XlsxFormulaEngine.fromFile(path)
        val sheet = "Three Statement Model"

        engine.setCellNumber(sheet, "D60", 2.0)
        val high = engine.evaluateCell(sheet, "D3")
        assertEquals(CellValue.TextValue("ERROR"), high)

        engine.setCellNumber(sheet, "D60", 0.5)
        val low = engine.evaluateCell(sheet, "D3")
        assertEquals(CellValue.TextValue("OK"), low)
    }

    @Test
    fun `financial model construction e26 should follow nested IF-MAX-SUM formula`() {
        val path = Path("src/test/resources/Financial+Model.xlsx")
        val engine = XlsxFormulaEngine.fromFile(path)
        val sheet = "Construction"

        val c22 = numeric(engine.evaluateCell(sheet, "C22"))
        val e19 = numeric(engine.evaluateCell(sheet, "E19"))

        val input1 = 0.0
        engine.setCellNumber(sheet, "D26", input1)
        val actual1 = numeric(engine.evaluateCell(sheet, "E26"))
        val expected1 = if (max(c22 - input1, 0.0) > e19) e19 else max(c22 - input1, 0.0)
        assertClose(expected1, actual1, "Construction!E26 case 1")

        val input2 = c22 + 1000.0
        engine.setCellNumber(sheet, "D26", input2)
        val actual2 = numeric(engine.evaluateCell(sheet, "E26"))
        val expected2 = if (max(c22 - input2, 0.0) > e19) e19 else max(c22 - input2, 0.0)
        assertClose(expected2, actual2, "Construction!E26 case 2")
    }

    @Test
    fun `statmnts income c15 should follow IF-MAX with cross-sheet refs`() {
        val path = Path("src/test/resources/statmnts.xlsx")
        val engine = XlsxFormulaEngine.fromFile(path)

        val opC9 = numeric(engine.evaluateCell("Operations", "C9"))
        val c12 = numeric(engine.evaluateCell("Income", "C12"))

        val b14Positive = 10.0
        engine.setCellNumber("Income", "B14", b14Positive)
        val positiveActual = numeric(engine.evaluateCell("Income", "C15"))
        val positiveExpected = max(opC9 * c12, -opC9 * b14Positive)
        assertClose(positiveExpected, positiveActual, "Income!C15 positive branch")

        val b14Negative = -5.0
        engine.setCellNumber("Income", "B14", b14Negative)
        val negativeActual = numeric(engine.evaluateCell("Income", "C15"))
        val negativeExpected = max(0.0, opC9 * (c12 + b14Negative))
        assertClose(negativeExpected, negativeActual, "Income!C15 negative branch")
    }

    private fun numeric(value: CellValue): Double {
        return (value as? CellValue.NumberValue)?.value
            ?: error("Expected NumberValue, got $value")
    }

    private fun assertClose(expected: Double, actual: Double, label: String) {
        val delta = abs(expected - actual)
        assertTrue(
            delta <= 1e-6,
            "$label mismatch. expected=$expected actual=$actual delta=$delta"
        )
    }
}
