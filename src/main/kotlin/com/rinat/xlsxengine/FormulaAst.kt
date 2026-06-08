package com.rinat.xlsxengine

sealed interface FormulaNode {
    data class NumberLiteral(val value: Double) : FormulaNode
    data class StringLiteral(val value: String) : FormulaNode
    data class BooleanLiteral(val value: Boolean) : FormulaNode
    data class ErrorLiteral(val code: String) : FormulaNode
    data object BlankLiteral : FormulaNode
    data class UnaryOp(val op: String, val expr: FormulaNode) : FormulaNode
    data class PostfixOp(val op: String, val expr: FormulaNode) : FormulaNode
    data class BinaryOp(val op: String, val left: FormulaNode, val right: FormulaNode) : FormulaNode
    data class CellRef(val sheet: String?, val ref: CellAddress) : FormulaNode
    data class RangeRef(val sheet: String?, val start: CellAddress, val end: CellAddress) : FormulaNode
    data class Range3DRef(
        val startSheet: String,
        val endSheet: String,
        val start: CellAddress,
        val end: CellAddress
    ) : FormulaNode
    data class DefinedNameRef(val name: String) : FormulaNode
    data class FunctionCall(val name: String, val args: List<FormulaNode>) : FormulaNode
}
