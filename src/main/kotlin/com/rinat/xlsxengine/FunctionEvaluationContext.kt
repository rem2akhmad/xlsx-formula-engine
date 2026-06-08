package com.rinat.xlsxengine

internal interface FunctionEvaluationContext {
    fun evalNode(currentSheet: String, node: FormulaNode): EvalValue
    fun evalScalar(currentSheet: String, node: FormulaNode): CellValue
    fun evaluateCell(sheet: String, address: CellAddress): CellValue
    fun iterateRange(sheet: String, start: CellAddress, end: CellAddress): Sequence<CellValue>
    fun forEachArgValue(
        currentSheet: String,
        args: List<FormulaNode>,
        onValue: (CellValue) -> Unit
    ): CellValue.ErrorValue?

    fun collectArgValues(currentSheet: String, args: List<FormulaNode>): List<CellValue>
    fun asBoolean(value: CellValue): Boolean?
}
