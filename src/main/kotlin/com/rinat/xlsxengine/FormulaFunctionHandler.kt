package com.rinat.xlsxengine

internal fun interface FormulaFunctionHandler {
    fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue
}
