package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage2Test {
    @Test
    fun `MATCH supports exact and approximate`() {
        val literals = mapOf(
            "A2" to num(10.0),
            "A3" to num(20.0),
            "A4" to num(30.0)
        )
        assertNumber(eval("MATCH(20,A2:A4,0)", literals), 2.0)
        assertNumber(eval("MATCH(25,A2:A4,1)", literals), 2.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.NA), eval("MATCH(5,A2:A4,1)", literals))
    }

    @Test
    fun `VLOOKUP and HLOOKUP`() {
        val literals = mapOf(
            "A2" to txt("A"), "B2" to num(100.0), "C2" to num(200.0),
            "A3" to txt("B"), "B3" to num(300.0), "C3" to num(400.0),
            "A4" to txt("C"), "B4" to num(500.0), "C4" to num(600.0),
            "A6" to num(1.0), "B6" to num(2.0), "C6" to num(3.0),
            "A7" to txt("x"), "B7" to txt("y"), "C7" to txt("z")
        )
        assertNumber(eval("VLOOKUP(\"B\",A2:C4,3,FALSE)", literals), 400.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.NA), eval("VLOOKUP(\"D\",A2:C4,2,FALSE)", literals))

        assertEquals(CellValue.TextValue("y"), eval("HLOOKUP(2,A6:C7,2,FALSE)", literals))
    }

    @Test
    fun `SUMIF COUNTIF AVERAGEIF`() {
        val literals = mapOf(
            "A2" to num(1.0), "A3" to num(2.0), "A4" to num(3.0), "A5" to num(4.0),
            "B2" to num(10.0), "B3" to num(20.0), "B4" to num(30.0), "B5" to num(40.0),
            "C2" to txt("aa"), "C3" to txt("ab"), "C4" to txt("ba"), "C5" to txt("bb")
        )
        assertNumber(eval("SUMIF(A2:A5,\">=3\",B2:B5)", literals), 70.0)
        assertNumber(eval("COUNTIF(C2:C5,\"a*\")", literals), 2.0)
        assertNumber(eval("AVERAGEIF(A2:A5,\">2\",B2:B5)", literals), 35.0)
    }

    @Test
    fun `OFFSET and INDIRECT references`() {
        val literals = mapOf(
            "A2" to num(5.0),
            "A3" to num(7.0),
            "B2" to num(11.0),
            "B3" to num(13.0)
        )
        assertNumber(eval("SUM(OFFSET(A2,0,0,2,1))", literals), 12.0)
        assertNumber(eval("SUM(INDIRECT(\"A2:A3\"))", literals), 12.0)
        assertNumber(eval("SUM(INDIRECT(\"B2:B3\"))", literals), 24.0)
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
