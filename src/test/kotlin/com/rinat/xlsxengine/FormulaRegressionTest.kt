package com.rinat.xlsxengine

import kotlin.io.path.Path
import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertTrue

class FormulaRegressionTest {
    @Test
    fun `sheet3 f4 in savings modeling should evaluate without parse error`() {
        val path = Path("src/test/resources/Savings modeling.xlsx")
        val workbook = XlsxReader.read(path)
        val sheet = workbook.sheets["Sheet3"] ?: error("Sheet3 not found")
        val cell = sheet.cells[CellAddress.parse("F4")] ?: error("Sheet3!F4 not found")
        val expected = cell.rawValue?.toDoubleOrNull() ?: error("Expected cached numeric value in Sheet3!F4")

        val engine = XlsxFormulaEngine.fromFile(path)
        val actual = engine.evaluateCell("Sheet3", "F4")

        assertTrue(actual is CellValue.NumberValue, "Expected numeric value for Sheet3!F4, got $actual")
        val delta = abs(expected - actual.value)
        assertTrue(delta <= 1e-9, "Mismatch for Sheet3!F4 expected=$expected actual=${actual.value} delta=$delta")
    }
}
