package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage3Test {
    @Test
    fun `SUMIFS COUNTIFS AVERAGEIFS multiple criteria`() {
        val literals = mapOf(
            "A2" to num(1.0), "A3" to num(1.0), "A4" to num(2.0), "A5" to num(2.0),
            "B2" to txt("x"), "B3" to txt("y"), "B4" to txt("x"), "B5" to txt("y"),
            "C2" to num(10.0), "C3" to num(20.0), "C4" to num(30.0), "C5" to num(40.0)
        )
        assertNumber(eval("SUMIFS(C2:C5,A2:A5,2,B2:B5,\"x\")", literals), 30.0)
        assertNumber(eval("COUNTIFS(A2:A5,1,B2:B5,\"y\")", literals), 1.0)
        assertNumber(eval("AVERAGEIFS(C2:C5,A2:A5,2,B2:B5,\"y\")", literals), 40.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.DIV_0), eval("AVERAGEIFS(C2:C5,A2:A5,3)", literals))
    }

    @Test
    fun `LOOKUP vector form`() {
        val literals = mapOf(
            "A2" to num(10.0), "A3" to num(20.0), "A4" to num(30.0),
            "B2" to txt("a"), "B3" to txt("b"), "B4" to txt("c")
        )
        assertEquals(CellValue.TextValue("b"), eval("LOOKUP(25,A2:A4,B2:B4)", literals))
        assertEquals(CellValue.ErrorValue(ErrorCodes.NA), eval("LOOKUP(5,A2:A4,B2:B4)", literals))
        assertNumber(eval("LOOKUP(25,A2:A4)", literals), 20.0)
    }

    @Test
    fun `EDATE and EOMONTH`() {
        assertNumber(eval("EDATE(45292,1)"), 45323.0)
        assertNumber(eval("EOMONTH(45292,0)"), 45322.0)
        assertNumber(eval("EOMONTH(45292,1)"), 45351.0)
    }

    @Test
    fun `NETWORKDAYS with and without holidays`() {
        val literals = mapOf("H1" to num(45294.0))
        assertNumber(eval("NETWORKDAYS(45292,45296)"), 5.0)
        assertNumber(eval("NETWORKDAYS(45292,45296,H1:H1)", literals), 4.0)
        assertNumber(eval("NETWORKDAYS(45296,45292)"), -5.0)
    }

    private fun eval(formula: String, literals: Map<String, CellValue> = emptyMap()): CellValue {
        val cells = linkedMapOf<CellAddress, WorkbookCell>()

        for ((a1, v) in literals) {
            val address = CellAddress.parse(a1)
            cells[address] = toWorkbookCell(address, v)
        }

        val a1 = CellAddress.parse("A1")
        cells[a1] = WorkbookCell(
            address = a1,
            type = null,
            rawValue = null,
            formula = formula
        )

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
    private fun txt(v: String): CellValue = CellValue.TextValue(v)

    private fun assertNumber(value: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(value is CellValue.NumberValue, "Expected number=$expected but got $value")
        assertTrue(abs(value.value - expected) <= eps, "Expected $expected but got ${value.value}")
    }
}
