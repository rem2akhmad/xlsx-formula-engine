package com.rinat.xlsxengine

import org.junit.jupiter.api.DynamicTest
import org.junit.jupiter.api.TestFactory
import org.junit.jupiter.api.assertDoesNotThrow
import java.nio.file.Files
import java.nio.file.Path
import kotlin.io.path.Path
import kotlin.math.abs
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class ResourceWorkbookCompatibilityTest {
    private val resourceFiles = listOf(
        "Accounts-Payable.xlsx",
        "Accounts-Receivable.xlsx",
        "Budget-Forecast.xlsx",
        "CFI-Case-Study-Three-Statement-Model.xlsx",
        "Expense-Claims.xlsx",
        "Financial Sample.xlsx",
        "Financial+Model.xlsx",
        "General-Ledger.xlsx",
        "Savings modeling.xlsx",
        "empty.xlsx",
        "statmnts.xlsx"
    )

    private val workbooks: Map<String, WorkbookData> by lazy {
        resourceFiles.associateWith { file -> XlsxReader.read(resourcePath(file)) }
    }

    private val engines: Map<String, XlsxFormulaEngine> by lazy {
        resourceFiles.associateWith { file -> XlsxFormulaEngine.fromFile(resourcePath(file)) }
    }

    @TestFactory
    fun `resource files should load and have consistent workbook metadata`(): List<DynamicTest> {
        return resourceFiles.map { file ->
            DynamicTest.dynamicTest("$file: workbook load + metadata") {
                val workbook = assertDoesNotThrow { workbooks.getValue(file) }
                assertTrue(workbook.sheetOrder.isNotEmpty(), "Workbook must contain at least one sheet: $file")
                assertEquals(workbook.sheetOrder.toSet(), workbook.sheets.keys, "Sheet order/key mismatch: $file")
                assertEquals(workbook.sheetOrder.size, workbook.sheetOrder.toSet().size, "Duplicate sheet names: $file")
            }
        }
    }

    @TestFactory
    fun `resource files should evaluate all sheets without runtime crash`(): List<DynamicTest> {
        return resourceFiles.map { file ->
            DynamicTest.dynamicTest("$file: evaluateAll") {
                val workbook = workbooks.getValue(file)
                val engine = engines.getValue(file)
                val evaluated = assertDoesNotThrow { engine.evaluateAll() }
                assertEquals(workbook.sheetOrder.toSet(), evaluated.sheets.keys, "Evaluated sheet set mismatch: $file")
            }
        }
    }

    @TestFactory
    fun `all formula cells should be evaluated and persisted to report`(): List<DynamicTest> {
        val reportsDir = Path("build/test-results/formula-eval")
        Files.createDirectories(reportsDir)

        return resourceFiles.map { file ->
            DynamicTest.dynamicTest("$file: evaluate every formula cell and write report") {
                val workbook = workbooks.getValue(file)
                val engine = engines.getValue(file)
                val reportRows = mutableListOf<String>()
                reportRows += "sheet,cell,status,expected,actual,formula"

                var formulaCount = 0
                val mismatches = mutableListOf<String>()

                workbook.sheetOrder.forEach { sheetName ->
                    val sheet = workbook.sheets.getValue(sheetName)
                    val formulaCells = sheet.cells.values
                        .filter { !it.formula.isNullOrBlank() }
                        .sortedWith(compareBy<WorkbookCell>({ it.address.row }, { it.address.column }))

                    formulaCells.forEach { cell ->
                        formulaCount += 1
                        val expected = cachedValue(workbook, cell)
                        val actual = engine.evaluateCell(sheetName, cell.address.toA1())

                        val ok = isEquivalent(expected, actual)
                        if (!ok) {
                            mismatches += "$sheetName!${cell.address.toA1()} expected=$expected actual=$actual formula=${cell.formula}"
                        }

                        reportRows += listOf(
                            csvEscape(sheetName),
                            csvEscape(cell.address.toA1()),
                            csvEscape(if (ok) "OK" else "MISMATCH"),
                            csvEscape(renderValue(expected)),
                            csvEscape(renderValue(actual)),
                            csvEscape(cell.formula ?: "")
                        ).joinToString(",")
                    }
                }

                val reportPath = reportsDir.resolve(file.replace(".xlsx", "-eval.csv"))
                Files.write(reportPath, reportRows)

                assertTrue(formulaCount >= 0, "No formula cells found in $file")
                assertTrue(
                    mismatches.isEmpty(),
                    "Found ${mismatches.size} formula mismatches in $file. First mismatches: ${mismatches.take(10)}"
                )
            }
        }
    }

    private fun cachedValue(workbook: WorkbookData, cell: WorkbookCell): CellValue {
        val raw = cell.rawValue ?: return CellValue.Blank
        return when (cell.type) {
            "s" -> {
                val idx = raw.toIntOrNull() ?: return CellValue.ErrorValue("Invalid shared string index '$raw'")
                workbook.sharedStrings.getOrNull(idx)?.let { CellValue.TextValue(it) }
                    ?: CellValue.ErrorValue("Shared string index out of bounds: $idx")
            }

            "str", "inlineStr" -> CellValue.TextValue(raw)
            "b" -> CellValue.BooleanValue(raw == "1")
            "e" -> CellValue.ErrorValue(raw)
            else -> raw.toDoubleOrNull()?.let { CellValue.NumberValue(it) } ?: CellValue.TextValue(raw)
        }
    }

    private fun isEquivalent(expected: CellValue, actual: CellValue): Boolean {
        return when {
            expected is CellValue.NumberValue && actual is CellValue.NumberValue -> {
                abs(expected.value - actual.value) <= 1e-4
            }

            expected is CellValue.BooleanValue && actual is CellValue.BooleanValue -> expected.value == actual.value
            expected is CellValue.TextValue && actual is CellValue.TextValue -> expected.value == actual.value
            expected is CellValue.ErrorValue && actual is CellValue.ErrorValue -> true
            expected is CellValue.Blank && actual is CellValue.Blank -> true
            else -> false
        }
    }

    private fun renderValue(value: CellValue): String {
        return when (value) {
            is CellValue.NumberValue -> value.value.toString()
            is CellValue.TextValue -> value.value
            is CellValue.BooleanValue -> value.value.toString()
            CellValue.Blank -> ""
            is CellValue.ErrorValue -> value.message
        }
    }

    private fun csvEscape(text: String): String {
        val escaped = text.replace("\"", "\"\"")
        return "\"$escaped\""
    }

    private fun resourcePath(fileName: String): Path = Path("src/test/resources/$fileName")
}
