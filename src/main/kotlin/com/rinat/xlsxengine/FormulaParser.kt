package com.rinat.xlsxengine

private class FormulaParser(formula: String) {
    private val tokens: List<Token> = FormulaTokenizer(formula).tokenize()
    private var pos: Int = 0

    fun parse(): FormulaNode {
        val expr = parseComparison()
        expect(TokenType.EOF)
        return expr
    }

    private fun parseComparison(): FormulaNode {
        var node = parseConcat()
        while (
            match(TokenType.GT) ||
            match(TokenType.LT) ||
            match(TokenType.GTE) ||
            match(TokenType.LTE) ||
            match(TokenType.EQ) ||
            match(TokenType.NEQ)
        ) {
            val op = previous().text
            val right = parseConcat()
            node = FormulaNode.BinaryOp(op, node, right)
        }
        return node
    }

    private fun parseConcat(): FormulaNode {
        var node = parseExpression()
        while (match(TokenType.AMP)) {
            val op = previous().text
            val right = parseExpression()
            node = FormulaNode.BinaryOp(op, node, right)
        }
        return node
    }

    private fun parseExpression(): FormulaNode {
        var node = parseTerm()
        while (match(TokenType.PLUS) || match(TokenType.MINUS)) {
            val op = previous().text
            val right = parseTerm()
            node = FormulaNode.BinaryOp(op, node, right)
        }
        return node
    }

    private fun parseTerm(): FormulaNode {
        var node = parsePower()
        while (match(TokenType.STAR) || match(TokenType.SLASH)) {
            val op = previous().text
            val right = parsePower()
            node = FormulaNode.BinaryOp(op, node, right)
        }
        return node
    }

    private fun parsePower(): FormulaNode {
        var node = parseUnary()
        while (match(TokenType.CARET)) {
            val op = previous().text
            val right = parseUnary()
            node = FormulaNode.BinaryOp(op, node, right)
        }
        return node
    }

    private fun parseUnary(): FormulaNode {
        if (match(TokenType.PLUS)) return FormulaNode.UnaryOp("+", parseUnary())
        if (match(TokenType.MINUS)) return FormulaNode.UnaryOp("-", parseUnary())
        return parsePostfix()
    }

    private fun parsePostfix(): FormulaNode {
        var node = parsePrimary()
        while (match(TokenType.PERCENT)) {
            node = FormulaNode.PostfixOp("%", node)
        }
        return node
    }

    private fun parsePrimary(): FormulaNode {
        if (match(TokenType.NUMBER)) {
            return FormulaNode.NumberLiteral(previous().text.toDouble())
        }

        if (match(TokenType.STRING)) {
            return FormulaNode.StringLiteral(previous().text)
        }

        if (match(TokenType.ERROR)) {
            return FormulaNode.ErrorLiteral(previous().text)
        }

        if (match(TokenType.CELL)) {
            return parseCellOrRange(sheet = null, firstCellToken = previous())
        }

        if (match(TokenType.SHEET)) {
            val firstSheet = previous().text
            if (match(TokenType.COLON)) {
                val secondSheet = when {
                    match(TokenType.SHEET) -> previous().text
                    match(TokenType.IDENT) -> previous().text
                    else -> error("Expected second sheet name in 3D reference")
                }
                expect(TokenType.BANG)
                val firstCell = expect(TokenType.CELL)
                val start = CellAddress.parse(firstCell.text)
                val end = if (match(TokenType.COLON)) CellAddress.parse(expect(TokenType.CELL).text) else start
                return FormulaNode.Range3DRef(firstSheet, secondSheet, start, end)
            }
            expect(TokenType.BANG)
            val cellToken = expect(TokenType.CELL)
            return parseCellOrRange(sheet = firstSheet, firstCellToken = cellToken)
        }

        if (match(TokenType.IDENT)) {
            val ident = previous().text

            if (match(TokenType.COLON)) {
                val secondSheet = when {
                    match(TokenType.SHEET) -> previous().text
                    match(TokenType.IDENT) -> previous().text
                    else -> error("Expected second sheet name in 3D reference")
                }
                expect(TokenType.BANG)
                val firstCell = expect(TokenType.CELL)
                val start = CellAddress.parse(firstCell.text)
                val end = if (match(TokenType.COLON)) CellAddress.parse(expect(TokenType.CELL).text) else start
                return FormulaNode.Range3DRef(ident, secondSheet, start, end)
            }

            if (match(TokenType.BANG)) {
                val cellToken = expect(TokenType.CELL)
                return parseCellOrRange(sheet = ident, firstCellToken = cellToken)
            }

            if (match(TokenType.LPAREN)) {
                val args = mutableListOf<FormulaNode>()
                if (!check(TokenType.RPAREN)) {
                    do {
                        if (check(TokenType.COMMA) || check(TokenType.RPAREN)) {
                            args += FormulaNode.BlankLiteral
                        } else {
                            args += parseComparison()
                        }
                    } while (match(TokenType.COMMA))
                }
                expect(TokenType.RPAREN)
                return FormulaNode.FunctionCall(ident.uppercase(), args)
            }

            if (ident.equals("TRUE", ignoreCase = true)) {
                return FormulaNode.BooleanLiteral(true)
            }
            if (ident.equals("FALSE", ignoreCase = true)) {
                return FormulaNode.BooleanLiteral(false)
            }

            return FormulaNode.DefinedNameRef(ident)
        }

        if (match(TokenType.LPAREN)) {
            val expr = parseComparison()
            expect(TokenType.RPAREN)
            return expr
        }

        error("Unexpected token ${peek().type} in formula")
    }

    private fun parseCellOrRange(sheet: String?, firstCellToken: Token): FormulaNode {
        val firstCell = CellAddress.parse(firstCellToken.text)
        if (match(TokenType.COLON)) {
            val endSheet = parseOptionalSheetPrefix()
            val end = expect(TokenType.CELL)
            val startSheet = sheet ?: endSheet
            val finalEndSheet = endSheet ?: sheet
            if (startSheet != null && finalEndSheet != null && startSheet != finalEndSheet) {
                return FormulaNode.Range3DRef(
                    startSheet = startSheet,
                    endSheet = finalEndSheet,
                    start = firstCell,
                    end = CellAddress.parse(end.text)
                )
            }
            return FormulaNode.RangeRef(
                sheet = startSheet,
                start = firstCell,
                end = CellAddress.parse(end.text)
            )
        }
        return FormulaNode.CellRef(sheet = sheet, ref = firstCell)
    }

    private fun parseOptionalSheetPrefix(): String? {
        if (match(TokenType.SHEET)) {
            val name = previous().text
            expect(TokenType.BANG)
            return name
        }
        if (match(TokenType.IDENT)) {
            val maybeName = previous().text
            if (match(TokenType.BANG)) {
                return maybeName
            }
            pos--
        }
        return null
    }

    private fun match(type: TokenType): Boolean {
        if (!check(type)) return false
        pos++
        return true
    }

    private fun check(type: TokenType): Boolean = peek().type == type

    private fun expect(type: TokenType): Token {
        if (check(type)) {
            pos++
            return previous()
        }
        error("Expected $type but found ${peek().type}")
    }

    private fun previous(): Token = tokens[pos - 1]

    private fun peek(): Token = tokens[pos]
}

internal fun parseFormula(formula: String): FormulaNode {
    val trimmed = formula.trim().removePrefix("=")
    return FormulaParser(trimmed).parse()
}
