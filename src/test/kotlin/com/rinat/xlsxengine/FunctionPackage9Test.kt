package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.math.exp
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage9Test {
    @Test
    fun `areas and statistical functions`() {
        assertNum(eval("AREAS(A1:B2,C3:D4)"), 2.0)
        assertNum(eval("AVEDEV(2,4,4,4,5,5,7,9)"), 1.5)
        assertNum(eval("MODE(1,2,2,3)"), 2.0)
        assertNum(eval("GEOMEAN(2,8)"), 4.0)
        assertNum(eval("HARMEAN(2,4)"), 2.6666666666666665, eps = 1e-12)

        val literals = mapOf(
            "A2" to num(1.0), "A3" to num(2.0), "A4" to num(3.0), "A5" to num(4.0),
            "B2" to num(2.0), "B3" to num(4.0), "B4" to num(6.0), "B5" to num(8.0)
        )
        assertNum(eval("CORREL(A2:A5,B2:B5)", literals), 1.0, eps = 1e-12)
        assertNum(eval("MAXA(TRUE,FALSE,\"x\",2)"), 2.0)
    }

    @Test
    fun `is family and byte text compatibility`() {
        assertEquals(CellValue.BooleanValue(true), eval("ISERR(#VALUE!)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISERR(#N/A)"))
        assertEquals(CellValue.BooleanValue(true), eval("ISLOGICAL(TRUE)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISLOGICAL(1)"))
        assertEquals(CellValue.BooleanValue(true), eval("ISNONTEXT(1)"))
        assertEquals(CellValue.BooleanValue(false), eval("ISNONTEXT(\"x\")"))
        assertEquals(CellValue.BooleanValue(true), eval("ISTEXT(\"x\")"))
        assertEquals(CellValue.BooleanValue(true), eval("ISEVEN(-2.9)"))
        assertEquals(CellValue.BooleanValue(true), eval("ISODD(3.1)"))

        assertEquals(CellValue.TextValue("ab"), eval("LEFTB(\"abcd\",2)"))
        assertNum(eval("LENB(\"abcd\")"), 4.0)
        assertEquals(CellValue.TextValue("bc"), eval("MIDB(\"abcd\",2,2)"))
    }

    @Test
    fun `datedif days360 and distribution`() {
        assertNum(eval("DATEDIF(DATE(2024,1,1),DATE(2024,1,31),\"D\")"), 30.0)
        assertNum(eval("DAYS360(DATE(2024,1,31),DATE(2024,2,28))"), 28.0)
        assertNum(eval("EFFECT(0.1,4)"), 0.103812890625, eps = 1e-12)
        assertNum(eval("EXPONDIST(1,2,TRUE)"), 1.0 - exp(-2.0), eps = 1e-12)
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
