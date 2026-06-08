package com.rinat.xlsxengine

import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage6Test {
    @Test
    fun `CONCATENATE UPPER LOWER TRIM`() {
        val literals = mapOf(
            "A2" to txt("Hello"),
            "A3" to txt(" "),
            "A4" to txt("World")
        )
        assertEquals(CellValue.TextValue("Hello World"), eval("CONCATENATE(A2:A4)", literals))
        assertEquals(CellValue.TextValue("HELLO"), eval("UPPER(\"hello\")"))
        assertEquals(CellValue.TextValue("world"), eval("LOWER(\"WORLD\")"))
        assertEquals(CellValue.TextValue("a b c"), eval("TRIM(\"  a   b   c  \")"))
    }

    @Test
    fun `REPT EXACT VALUE`() {
        assertEquals(CellValue.TextValue("ababab"), eval("REPT(\"ab\",3)"))
        assertEquals(CellValue.BooleanValue(true), eval("EXACT(\"Abc\",\"Abc\")"))
        assertEquals(CellValue.BooleanValue(false), eval("EXACT(\"Abc\",\"abc\")"))
        assertNum(eval("VALUE(\" 123.5 \")"), 123.5)
        assertEquals(CellValue.ErrorValue(ErrorCodes.VALUE), eval("VALUE(\"abc\")"))
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

    private fun txt(v: String): CellValue = CellValue.TextValue(v)

    private fun assertNum(actual: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(actual is CellValue.NumberValue, "Expected number=$expected but got $actual")
        assertTrue(abs(actual.value - expected) <= eps, "Expected $expected but got ${actual.value}")
    }
}
