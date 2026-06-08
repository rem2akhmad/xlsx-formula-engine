package com.rinat.xlsxengine

import java.time.LocalDate
import java.time.temporal.ChronoUnit
import kotlin.math.abs
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class FunctionPackage8Test {
    @Test
    fun `text and counting functions`() {
        val literals = mapOf(
            "A2" to num(10.0),
            "A3" to txt("x"),
            "A4" to CellValue.Blank,
            "A5" to txt("")
        )
        assertEquals(CellValue.TextValue("A"), eval("CHAR(65)"))
        assertNum(eval("CODE(\"A\")"), 65.0)
        assertEquals(CellValue.TextValue("AB"), eval("CLEAN(B1)", mapOf("B1" to txt("A\u0001B"))))
        assertEquals(CellValue.TextValue("Hello World"), eval("PROPER(\"hello world\")"))
        assertEquals(CellValue.TextValue("abZZgh"), eval("REPLACE(\"abcdefgh\",3,4,\"ZZ\")"))

        assertNum(eval("COUNTA(A2:A5)", literals), 3.0)
        assertNum(eval("COUNTBLANK(A2:A5)", literals), 2.0)
    }

    @Test
    fun `number formatting and math helpers`() {
        assertEquals(CellValue.TextValue("\$1,234.57"), eval("DOLLAR(1234.567,2)"))
        assertEquals(CellValue.TextValue("1234.6"), eval("FIXED(1234.567,1,TRUE)"))
        assertNum(eval("COMBIN(5,2)"), 10.0)
        assertNum(eval("GCD(48,18)"), 6.0)
        assertNum(eval("GESTEP(5,4)"), 1.0)
        assertNum(eval("GESTEP(3,4)"), 0.0)
    }

    @Test
    fun `date-time and volatile`() {
        assertNum(eval("TIME(1,2,3)"), (3600.0 + 120.0 + 3.0) / 86400.0, eps = 1e-12)
        assertNum(eval("TIMEVALUE(\"13:05:09\")"), (13 * 3600 + 5 * 60 + 9) / 86400.0, eps = 1e-12)
        assertNum(eval("HOUR(TIME(13,30,0))"), 13.0)
        assertNum(eval("SECOND(TIME(13,30,45))"), 45.0)
        assertEquals(CellValue.TextValue("1234.57"), eval("TEXT(1234.567,\"0.00\")"))
        assertEquals(CellValue.TextValue("2024-01-01"), eval("TEXT(45292,\"yyyy-mm-dd\")"))

        val today = eval("TODAY()")
        assertTrue(today is CellValue.NumberValue)
        val expectedToday = ChronoUnit.DAYS.between(LocalDate.of(1899, 12, 30), LocalDate.now()).toDouble()
        assertTrue(abs(today.value - expectedToday) <= 1e-9, "Expected $expectedToday but got ${today.value}")

        val now = eval("NOW()")
        assertTrue(now is CellValue.NumberValue)
        assertTrue(now.value >= expectedToday && now.value < expectedToday + 1.0, "NOW out of expected day range: ${now.value}")

        val rand = eval("RAND()")
        assertTrue(rand is CellValue.NumberValue)
        assertTrue(rand.value >= 0.0 && rand.value < 1.0, "RAND out of range: ${rand.value}")
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
    private fun txt(v: String): CellValue = CellValue.TextValue(v)

    private fun assertNum(actual: CellValue, expected: Double, eps: Double = 1e-9) {
        assertTrue(actual is CellValue.NumberValue, "Expected number=$expected but got $actual")
        assertTrue(abs(actual.value - expected) <= eps, "Expected $expected but got ${actual.value}")
    }
}
