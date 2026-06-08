package com.rinat.xlsxengine

import kotlin.math.PI
import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage7Test {
    @Test
    fun `inverse trig and hyperbolic`() {
        assertNum(eval("ACOS(1)"), 0.0)
        assertNum(eval("ACOSH(1)"), 0.0)
        assertNum(eval("ASIN(1)"), PI / 2.0, eps = 1e-12)
        assertNum(eval("ASINH(0)"), 0.0)
        assertNum(eval("ATAN2(1,1)"), PI / 4.0, eps = 1e-12)
        assertNum(eval("ATANH(0.5)"), 0.5493061443340548, eps = 1e-12)
    }

    @Test
    fun `trig family and radians`() {
        assertNum(eval("COSH(0)"), 1.0)
        assertNum(eval("RADIANS(180)"), PI, eps = 1e-12)
        assertNum(eval("SIN(PI()/2)"), 1.0, eps = 1e-12)
        assertNum(eval("SINH(0)"), 0.0)
        assertNum(eval("TAN(0)"), 0.0)
        assertNum(eval("TANH(0)"), 0.0)
        assertNum(eval("SQRTPI(4)"), kotlin.math.sqrt(4.0 * PI), eps = 1e-12)
    }

    @Test
    fun `trunc quotient boolean and ref shape`() {
        assertNum(eval("TRUNC(8.9)"), 8.0)
        assertNum(eval("TRUNC(-8.9)"), -8.0)
        assertNum(eval("TRUNC(123.456,2)"), 123.45)
        assertNum(eval("TRUNC(123.456,-1)"), 120.0)

        assertNum(eval("QUOTIENT(5,2)"), 2.0)
        assertNum(eval("QUOTIENT(-5,2)"), -2.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.DIV_0), eval("QUOTIENT(5,0)"))

        assertEquals(CellValue.BooleanValue(true), eval("TRUE()"))
        assertEquals(CellValue.BooleanValue(false), eval("FALSE()"))

        assertNum(eval("COLUMN(B5)"), 2.0)
        assertNum(eval("COLUMN(B5:D9)"), 2.0)
        assertNum(eval("COLUMNS(B5:D9)"), 3.0)
        assertNum(eval("ROWS(B5:D9)"), 5.0)
    }

    private fun eval(formula: String, literals: Map<String, CellValue> = emptyMap()): CellValue {
        val cells = linkedMapOf<CellAddress, WorkbookCell>()
        for ((a1, v) in literals) {
            val address = CellAddress.parse(a1)
            cells[address] = toWorkbookCell(address, v)
        }

        val a1 = CellAddress.parse("A1")
        cells[a1] = WorkbookCell(a1, null, null, formula)

        val workbook = WorkbookData(
            sheets = mapOf("S" to WorksheetData("S", cells)),
            sheetOrder = listOf("S"),
            sharedStrings = emptyList()
        )
        val evaluator = DefaultFormulaEvaluator(workbook)
        return evaluator.evaluateCell("S", a1)
    }

    private fun toWorkbookCell(address: CellAddress, value: CellValue): WorkbookCell {
        return when (value) {
            is CellValue.NumberValue -> WorkbookCell(address, null, value.value.toString(), null)
            is CellValue.TextValue -> WorkbookCell(address, "str", value.value, null)
            is CellValue.BooleanValue -> WorkbookCell(address, "b", if (value.value) "1" else "0", null)
            CellValue.Blank -> WorkbookCell(address, null, null, null)
            is CellValue.ErrorValue -> WorkbookCell(address, "e", value.message, null)
        }
    }

    private fun assertNum(actual: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(actual is CellValue.NumberValue, "Expected number=$expected but got $actual")
        assertTrue(abs(actual.value - expected) <= eps, "Expected $expected but got ${actual.value}")
    }
}
