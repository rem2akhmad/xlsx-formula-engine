package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertTrue

class FunctionPackage4Test {
    @Test
    fun `RANK PERCENTILE QUARTILE`() {
        val literals = mapOf(
            "A2" to num(10.0), "A3" to num(20.0), "A4" to num(30.0), "A5" to num(40.0)
        )
        assertNum(eval("RANK(30,A2:A5)", literals), 2.0)
        assertNum(eval("RANK(30,A2:A5,1)", literals), 3.0)
        assertNum(eval("PERCENTILE(A2:A5,0.5)", literals), 25.0)
        assertNum(eval("QUARTILE(A2:A5,1)", literals), 17.5)
    }

    @Test
    fun `STDEV STDEVP VAR VARP`() {
        val literals = mapOf(
            "A2" to num(2.0), "A3" to num(4.0), "A4" to num(4.0), "A5" to num(4.0),
            "A6" to num(5.0), "A7" to num(5.0), "A8" to num(7.0), "A9" to num(9.0)
        )
        assertNum(eval("STDEV(A2:A9)", literals), 2.138089935299395, eps = 1e-9)
        assertNum(eval("STDEVP(A2:A9)", literals), 2.0, eps = 1e-9)
        assertNum(eval("VAR(A2:A9)", literals), 4.571428571428571, eps = 1e-9)
        assertNum(eval("VARP(A2:A9)", literals), 4.0, eps = 1e-9)
    }

    @Test
    fun `SUMX family`() {
        val literals = mapOf(
            "A2" to num(1.0), "A3" to num(2.0), "A4" to num(3.0),
            "B2" to num(4.0), "B3" to num(5.0), "B4" to num(6.0)
        )
        assertNum(eval("SUMX2MY2(A2:A4,B2:B4)", literals), -63.0)
        assertNum(eval("SUMX2PY2(A2:A4,B2:B4)", literals), 91.0)
        assertNum(eval("SUMXMY2(A2:A4,B2:B4)", literals), 27.0)
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
