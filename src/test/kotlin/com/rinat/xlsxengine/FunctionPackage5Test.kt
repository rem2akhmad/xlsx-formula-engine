package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.math.PI
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage5Test {
    @Test
    fun `MEDIAN SMALL LARGE`() {
        val literals = mapOf(
            "A2" to num(10.0), "A3" to num(20.0), "A4" to num(30.0), "A5" to num(40.0)
        )
        assertNum(eval("MEDIAN(A2:A5)", literals), 25.0)
        assertNum(eval("SMALL(A2:A5,2)", literals), 20.0)
        assertNum(eval("LARGE(A2:A5,3)", literals), 20.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.NUM), eval("SMALL(A2:A5,10)", literals))
    }

    @Test
    fun `PRODUCT and SUMSQ`() {
        val literals = mapOf(
            "A2" to num(4.0), "A3" to num(5.0)
        )
        assertNum(eval("PRODUCT(2,3,A2:A3)", literals), 120.0)
        assertNum(eval("SUMSQ(2,3,A2:A3)", literals), 54.0)
    }

    @Test
    fun `CEILING FLOOR MROUND SIGN PI`() {
        assertNum(eval("CEILING(2.1)"), 3.0)
        assertNum(eval("CEILING(10,3)"), 12.0)
        assertNum(eval("FLOOR(10,3)"), 9.0)
        assertNum(eval("MROUND(10,3)"), 9.0)
        assertNum(eval("MROUND(11,3)"), 12.0)
        assertNum(eval("SIGN(-10)"), -1.0)
        assertNum(eval("SIGN(0)"), 0.0)
        assertNum(eval("SIGN(10)"), 1.0)
        assertNum(eval("PI()"), PI, eps = 1e-12)
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

    private fun num(v: Double): CellValue = CellValue.NumberValue(v)

    private fun assertNum(actual: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(actual is CellValue.NumberValue, "Expected number=$expected but got $actual")
        assertTrue(abs(actual.value - expected) <= eps, "Expected $expected but got ${actual.value}")
    }
}
