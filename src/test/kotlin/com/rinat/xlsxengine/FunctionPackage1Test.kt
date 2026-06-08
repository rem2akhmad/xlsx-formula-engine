package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage1Test {
    @Test
    fun `IFNA and IS family`() {
        assertEquals(CellValue.NumberValue(5.0), eval("IFNA(#N/A,5)"))
        assertEquals(CellValue.ErrorValue(ErrorCodes.VALUE), eval("IFNA(#VALUE!,5)"))

        assertEquals(CellValue.BooleanValue(true), eval("ISERROR(#REF!)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISERROR(1)"))
        assertEquals(CellValue.BooleanValue(true), eval("ISNA(#N/A)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISNA(#REF!)"))
        assertEquals(CellValue.BooleanValue(true), eval("ISNUMBER(12.3)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISNUMBER(\"x\")"))
        assertEquals(CellValue.BooleanValue(true), eval("ISBLANK(B1)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISBLANK(0)"))
    }

    @Test
    fun `AND OR NOT and CHOOSE`() {
        assertEquals(CellValue.BooleanValue(true), eval("AND(TRUE,1,2)"))
        assertEquals(CellValue.BooleanValue(false), eval("AND(TRUE,0)"))
        assertEquals(CellValue.BooleanValue(true), eval("OR(FALSE,0,1)"))
        assertEquals(CellValue.BooleanValue(false), eval("OR(FALSE,0)"))
        assertEquals(CellValue.BooleanValue(false), eval("NOT(TRUE)"))

        assertEquals(CellValue.TextValue("b"), eval("CHOOSE(2,\"a\",\"b\",\"c\")"))
        assertEquals(CellValue.ErrorValue(ErrorCodes.VALUE), eval("CHOOSE(5,1,2)"))
    }

    @Test
    fun `POWER INT MOD ROUND family`() {
        assertNumber(eval("POWER(2,3)"), 8.0)
        assertNumber(eval("INT(1.9)"), 1.0)
        assertNumber(eval("INT(-1.9)"), -2.0)
        assertNumber(eval("MOD(10,3)"), 1.0)
        assertNumber(eval("MOD(-3,2)"), 1.0)

        assertNumber(eval("ROUND(1.25,1)"), 1.3)
        assertNumber(eval("ROUND(-1.25,1)"), -1.3)
        assertNumber(eval("ROUNDUP(1.21,1)"), 1.3)
        assertNumber(eval("ROUNDUP(-1.21,1)"), -1.3)
        assertNumber(eval("ROUNDDOWN(1.29,1)"), 1.2)
        assertNumber(eval("ROUNDDOWN(-1.29,1)"), -1.2)
    }

    @Test
    fun `text functions`() {
        assertNumber(eval("LEN(\"abcd\")"), 4.0)
        assertEquals(CellValue.TextValue("ab"), eval("LEFT(\"abcd\",2)"))
        assertEquals(CellValue.TextValue("cd"), eval("RIGHT(\"abcd\",2)"))
        assertEquals(CellValue.TextValue("bc"), eval("MID(\"abcd\",2,2)"))

        assertNumber(eval("FIND(\"BC\",\"ABCD\")"), 2.0)
        assertEquals(CellValue.ErrorValue(ErrorCodes.VALUE), eval("FIND(\"BC\",\"ABCD\",3)"))
        assertNumber(eval("SEARCH(\"bc\",\"ABCD\")"), 2.0)
        assertEquals(CellValue.TextValue("a-b+c"), eval("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)"))
    }

    @Test
    fun `date functions`() {
        assertNumber(eval("DATE(2024,1,1)"), 45292.0)
        assertNumber(eval("YEAR(45292)"), 2024.0)
        assertNumber(eval("MONTH(45292)"), 1.0)
        assertNumber(eval("DAY(45292)"), 1.0)
    }

    @Test
    fun `SUMPRODUCT works for vectors`() {
        val literals = mapOf(
            "A2" to num(2.0),
            "A3" to num(3.0),
            "B2" to num(4.0),
            "B3" to num(5.0)
        )
        assertNumber(eval("SUMPRODUCT(A2:A3,B2:B3)", literals), 23.0)
        assertNumber(eval("SUMPRODUCT(A2:A3,2)", literals), 10.0)
    }

    @Test
    fun `financial functions PMT PV FV NPER RATE`() {
        val pmt = eval("PMT(0.05,12,1000)")
        assertTrue(pmt is CellValue.NumberValue)

        val literals = mapOf("B1" to lit(pmt))
        assertNumber(eval("PV(0.05,12,B1)", literals), 1000.0, eps = 1e-6)

        assertNumber(eval("FV(0.1,2,-100,0,0)"), 210.0, eps = 1e-6)
        assertNumber(eval("NPER(0.1,-200,1000,0,0)"), 7.272540897341713, eps = 1e-6)

        val rateWorkbook = mapOf(
            "B1" to lit(eval("PMT(0.05,12,1000)"))
        )
        assertNumber(eval("RATE(12,B1,1000,0,0)", rateWorkbook), 0.05, eps = 1e-6)
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

    private fun lit(v: CellValue): CellValue = v

    private fun assertNumber(value: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(value is CellValue.NumberValue, "Expected number=$expected but got $value")
        assertTrue(abs(value.value - expected) <= eps, "Expected $expected but got ${value.value}")
    }
}
