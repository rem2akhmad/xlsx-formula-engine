package com.rinat.xlsxengine

import org.junit.jupiter.api.DynamicTest
import org.junit.jupiter.api.TestFactory
import java.nio.file.Path
import kotlin.io.path.Path
import kotlin.math.abs
import kotlin.test.assertTrue

class WorkbookMutationRecalculationTest {
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

    @TestFactory
    fun `changing an input cell should recalculate by formula semantics`(): List<DynamicTest> {
        return resourceFiles.map { file ->
            DynamicTest.dynamicTest("$file: mutate input and verify exact formula recalculation") {
                val workbook = XlsxReader.read(resourcePath(file))
                val formulaCount = workbook.sheets.values.sumOf { sheet ->
                    sheet.cells.values.count { !it.formula.isNullOrBlank() }
                }

                if (formulaCount == 0) {
                    assertTrue(true, "No formulas in $file")
                    return@dynamicTest
                }

                val scenario = findExactRecalculationScenario(file, workbook)
                    ?: error("No exact linear recalculation scenario found in $file")

                val diff = abs(scenario.afterTarget - scenario.expectedAfter)
                assertTrue(
                    diff <= 1e-6,
                    "Expected exact recalculation in $file for ${scenario.targetSheet}!${scenario.targetCell}. " +
                        "formula='${scenario.formula}' input=${scenario.inputSheet}!${scenario.inputCell} " +
                        "before=${scenario.beforeTarget} expected=${scenario.expectedAfter} actual=${scenario.afterTarget} diff=$diff"
                )
            }
        }
    }

    private data class Scenario(
        val targetSheet: String,
        val targetCell: String,
        val inputSheet: String,
        val inputCell: String,
        val formula: String,
        val beforeTarget: Double,
        val expectedAfter: Double,
        val afterTarget: Double
    )

    private data class RefCell(val sheet: String, val address: CellAddress)
    private data class LinearExpr(val coef: Double, val constant: Double)

    private fun findExactRecalculationScenario(fileName: String, workbook: WorkbookData): Scenario? {
        val baseEngine = XlsxFormulaEngine.fromFile(resourcePath(fileName))
        val mutationDelta = 1234.567

        for (sheetName in workbook.sheetOrder) {
            val sheet = workbook.sheets.getValue(sheetName)
            val formulaCells = sheet.cells.values
                .filter { !it.formula.isNullOrBlank() }
                .sortedWith(compareBy<WorkbookCell>({ it.address.row }, { it.address.column }))

            for (cell in formulaCells) {
                val formula = cell.formula ?: continue
                val ast = runCatching { parseFormula(formula) }.getOrNull() ?: continue

                val before = baseEngine.evaluateCell(sheetName, cell.address.toA1())
                val beforeTarget = (before as? CellValue.NumberValue)?.value ?: continue

                val refs = collectReferencedCells(ast, sheetName)
                for (ref in refs) {
                    if (ref.sheet == sheetName && ref.address == cell.address) continue

                    val refSheet = workbook.sheets[ref.sheet] ?: continue
                    val refCell = refSheet.cells[ref.address] ?: continue
                    if (!refCell.formula.isNullOrBlank()) continue

                    val inputValue = baseEngine.evaluateCell(ref.sheet, ref.address.toA1())
                    val inputNumber = (inputValue as? CellValue.NumberValue)?.value ?: continue

                    val expr = linearExpr(ast, sheetName, ref, baseEngine)
                    if (expr == null || abs(expr.coef) < 1e-12) continue

                    val engine = XlsxFormulaEngine.fromFile(resourcePath(fileName))
                    engine.setCellNumber(ref.sheet, ref.address.toA1(), inputNumber + mutationDelta)
                    val updated = engine.evaluateCell(sheetName, cell.address.toA1())
                    val afterTarget = (updated as? CellValue.NumberValue)?.value ?: continue

                    val expectedAfter = beforeTarget + expr.coef * mutationDelta
                    return Scenario(
                        targetSheet = sheetName,
                        targetCell = cell.address.toA1(),
                        inputSheet = ref.sheet,
                        inputCell = ref.address.toA1(),
                        formula = formula,
                        beforeTarget = beforeTarget,
                        expectedAfter = expectedAfter,
                        afterTarget = afterTarget
                    )
                }
            }
        }

        return null
    }

    private fun collectReferencedCells(node: FormulaNode, currentSheet: String): Set<RefCell> {
        return when (node) {
            is FormulaNode.CellRef -> setOf(RefCell(node.sheet ?: currentSheet, node.ref))
            is FormulaNode.RangeRef -> {
                val sheet = node.sheet ?: currentSheet
                val refs = mutableSetOf<RefCell>()
                val rowStart = minOf(node.start.row, node.end.row)
                val rowEnd = maxOf(node.start.row, node.end.row)
                val colStart = minOf(node.start.column, node.end.column)
                val colEnd = maxOf(node.start.column, node.end.column)
                for (row in rowStart..rowEnd) {
                    for (col in colStart..colEnd) {
                        refs += RefCell(sheet, CellAddress(row, col))
                    }
                }
                refs
            }

            is FormulaNode.Range3DRef -> emptySet()

            is FormulaNode.BinaryOp -> collectReferencedCells(node.left, currentSheet) +
                collectReferencedCells(node.right, currentSheet)

            is FormulaNode.UnaryOp -> collectReferencedCells(node.expr, currentSheet)
            is FormulaNode.PostfixOp -> collectReferencedCells(node.expr, currentSheet)
            is FormulaNode.FunctionCall -> node.args.flatMapTo(mutableSetOf()) { collectReferencedCells(it, currentSheet) }
            is FormulaNode.DefinedNameRef,
            is FormulaNode.ErrorLiteral,
            is FormulaNode.NumberLiteral,
            is FormulaNode.StringLiteral,
            is FormulaNode.BooleanLiteral,
            FormulaNode.BlankLiteral -> emptySet()
        }
    }

    private fun linearExpr(
        node: FormulaNode,
        currentSheet: String,
        ref: RefCell,
        engine: XlsxFormulaEngine
    ): LinearExpr? {
        return when (node) {
            is FormulaNode.NumberLiteral -> LinearExpr(0.0, node.value)
            FormulaNode.BlankLiteral -> LinearExpr(0.0, 0.0)
            is FormulaNode.BooleanLiteral -> LinearExpr(0.0, if (node.value) 1.0 else 0.0)
            is FormulaNode.StringLiteral -> node.value.toDoubleOrNull()?.let { LinearExpr(0.0, it) } ?: null

            is FormulaNode.CellRef -> {
                val sheet = node.sheet ?: currentSheet
                if (sheet == ref.sheet && node.ref == ref.address) {
                    LinearExpr(1.0, 0.0)
                } else {
                    val value = engine.evaluateCell(sheet, node.ref.toA1()).asNumberOrNull() ?: return null
                    LinearExpr(0.0, value)
                }
            }

            is FormulaNode.RangeRef -> {
                val sheet = node.sheet ?: currentSheet
                val rowStart = minOf(node.start.row, node.end.row)
                val rowEnd = maxOf(node.start.row, node.end.row)
                val colStart = minOf(node.start.column, node.end.column)
                val colEnd = maxOf(node.start.column, node.end.column)

                var coef = 0.0
                var constant = 0.0
                for (row in rowStart..rowEnd) {
                    for (col in colStart..colEnd) {
                        val address = CellAddress(row, col)
                        if (sheet == ref.sheet && address == ref.address) {
                            coef += 1.0
                        } else {
                            val value = engine.evaluateCell(sheet, address.toA1())
                            constant += value.asNumberOrNull() ?: 0.0
                        }
                    }
                }
                LinearExpr(coef, constant)
            }

            is FormulaNode.Range3DRef -> null

            is FormulaNode.UnaryOp -> {
                val c = linearExpr(node.expr, currentSheet, ref, engine) ?: return null
                when (node.op) {
                    "+" -> c
                    "-" -> LinearExpr(-c.coef, -c.constant)
                    else -> null
                }
            }

            is FormulaNode.PostfixOp -> {
                if (node.op != "%") return null
                val c = linearExpr(node.expr, currentSheet, ref, engine) ?: return null
                LinearExpr(c.coef / 100.0, c.constant / 100.0)
            }

            is FormulaNode.BinaryOp -> {
                val left = linearExpr(node.left, currentSheet, ref, engine) ?: return null
                val right = linearExpr(node.right, currentSheet, ref, engine) ?: return null
                when (node.op) {
                    "+" -> LinearExpr(left.coef + right.coef, left.constant + right.constant)
                    "-" -> LinearExpr(left.coef - right.coef, left.constant - right.constant)
                    "*" -> {
                        if (abs(left.coef) > 1e-12 && abs(right.coef) > 1e-12) return null
                        val coef = left.coef * right.constant + right.coef * left.constant
                        val constant = left.constant * right.constant
                        LinearExpr(coef, constant)
                    }

                    "/" -> {
                        if (abs(right.coef) > 1e-12) return null
                        if (abs(right.constant) <= 1e-12) return null
                        LinearExpr(left.coef / right.constant, left.constant / right.constant)
                    }

                    else -> null
                }
            }

            is FormulaNode.FunctionCall -> {
                if (node.name != "SUM") return null
                var totalCoef = 0.0
                var totalConst = 0.0
                for (arg in node.args) {
                    val c = linearExpr(arg, currentSheet, ref, engine) ?: return null
                    totalCoef += c.coef
                    totalConst += c.constant
                }
                LinearExpr(totalCoef, totalConst)
            }

            is FormulaNode.DefinedNameRef,
            is FormulaNode.ErrorLiteral -> null
        }
    }

    private fun resourcePath(fileName: String): Path = Path("src/test/resources/$fileName")
}
