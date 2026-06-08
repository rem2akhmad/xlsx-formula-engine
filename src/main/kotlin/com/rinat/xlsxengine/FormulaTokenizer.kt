package com.rinat.xlsxengine

internal enum class TokenType {
    NUMBER,
    STRING,
    ERROR,
    IDENT,
    SHEET,
    CELL,
    BANG,
    PLUS,
    MINUS,
    STAR,
    SLASH,
    CARET,
    PERCENT,
    AMP,
    GT,
    LT,
    GTE,
    LTE,
    EQ,
    NEQ,
    COLON,
    COMMA,
    LPAREN,
    RPAREN,
    EOF
}

internal data class Token(val type: TokenType, val text: String)

internal class FormulaTokenizer(private val source: String) {
    private var index = 0

    fun tokenize(): List<Token> {
        val out = mutableListOf<Token>()
        while (true) {
            skipSpaces()
            if (index >= source.length) {
                out += Token(TokenType.EOF, "")
                return out
            }

            val ch = source[index]
            when {
                ch.isDigit() || (ch == '.' && peek(1)?.isDigit() == true) -> out += readNumber()
                ch == '"' -> out += readString()
                ch == '\'' -> out += readQuotedSheetName()

                ch == '>' && peek(1) == '=' -> {
                    out += Token(TokenType.GTE, ">=")
                    index += 2
                }

                ch == '<' && peek(1) == '=' -> {
                    out += Token(TokenType.LTE, "<=")
                    index += 2
                }

                ch == '<' && peek(1) == '>' -> {
                    out += Token(TokenType.NEQ, "<>")
                    index += 2
                }

                ch == '+' -> {
                    out += Token(TokenType.PLUS, "+")
                    index++
                }

                ch == '-' -> {
                    out += Token(TokenType.MINUS, "-")
                    index++
                }

                ch == '*' -> {
                    out += Token(TokenType.STAR, "*")
                    index++
                }

                ch == '/' -> {
                    out += Token(TokenType.SLASH, "/")
                    index++
                }

                ch == '^' -> {
                    out += Token(TokenType.CARET, "^")
                    index++
                }

                ch == '%' -> {
                    out += Token(TokenType.PERCENT, "%")
                    index++
                }

                ch == '&' -> {
                    out += Token(TokenType.AMP, "&")
                    index++
                }

                ch == '>' -> {
                    out += Token(TokenType.GT, ">")
                    index++
                }

                ch == '<' -> {
                    out += Token(TokenType.LT, "<")
                    index++
                }

                ch == '=' -> {
                    out += Token(TokenType.EQ, "=")
                    index++
                }

                ch == ':' -> {
                    out += Token(TokenType.COLON, ":")
                    index++
                }

                ch == ',' -> {
                    out += Token(TokenType.COMMA, ",")
                    index++
                }

                ch == '(' -> {
                    out += Token(TokenType.LPAREN, "(")
                    index++
                }

                ch == ')' -> {
                    out += Token(TokenType.RPAREN, ")")
                    index++
                }

                ch == '!' -> {
                    out += Token(TokenType.BANG, "!")
                    index++
                }

                ch == '$' || ch.isLetter() || ch == '_' -> out += readWordLike()
                ch == '#' -> out += readErrorLiteral()
                else -> error("Unsupported formula token '$ch' in '$source'")
            }
        }
    }

    private fun readNumber(): Token {
        val start = index
        var sawDot = false
        while (index < source.length) {
            val ch = source[index]
            when {
                ch.isDigit() -> index++
                ch == '.' && !sawDot -> {
                    sawDot = true
                    index++
                }

                else -> break
            }
        }
        return Token(TokenType.NUMBER, source.substring(start, index))
    }

    private fun readString(): Token {
        index++
        val sb = StringBuilder()
        while (index < source.length) {
            val ch = source[index]
            index++
            if (ch == '"') {
                if (index < source.length && source[index] == '"') {
                    sb.append('"')
                    index++
                    continue
                }
                return Token(TokenType.STRING, sb.toString())
            }
            sb.append(ch)
        }
        error("Unterminated string in formula '$source'")
    }

    private fun readQuotedSheetName(): Token {
        index++
        val sb = StringBuilder()
        while (index < source.length) {
            val ch = source[index]
            index++
            if (ch == '\'') {
                if (index < source.length && source[index] == '\'') {
                    sb.append('\'')
                    index++
                    continue
                }
                return Token(TokenType.SHEET, sb.toString())
            }
            sb.append(ch)
        }
        error("Unterminated quoted sheet name in formula '$source'")
    }

    private fun readWordLike(): Token {
        val start = index
        while (index < source.length) {
            val ch = source[index]
            if (ch.isLetterOrDigit() || ch == '_' || ch == '$' || ch == '.') {
                index++
            } else {
                break
            }
        }
        val text = source.substring(start, index)
        val next = nextNonWhitespaceFrom(index)
        val callLike = next == '('
        return if (isCellReference(text) && !callLike) Token(TokenType.CELL, text) else Token(TokenType.IDENT, text)
    }

    private fun readErrorLiteral(): Token {
        val start = index
        index++
        while (index < source.length) {
            val ch = source[index]
            if (ch.isLetterOrDigit() || ch == '/' || ch == '!' || ch == '?' || ch == '#') {
                index++
            } else {
                break
            }
        }
        val text = source.substring(start, index).uppercase()
        return Token(TokenType.ERROR, text)
    }

    private fun skipSpaces() {
        while (index < source.length && source[index].isWhitespace()) {
            index++
        }
    }

    private fun peek(offset: Int): Char? {
        val i = index + offset
        return if (i in source.indices) source[i] else null
    }

    private fun nextNonWhitespaceFrom(from: Int): Char? {
        var i = from
        while (i < source.length) {
            if (!source[i].isWhitespace()) return source[i]
            i++
        }
        return null
    }

    private fun isCellReference(raw: String): Boolean {
        if (R1C1_REGEX.matches(raw.uppercase())) return true

        val text = raw.replace("$", "")
        val splitIndex = text.indexOfFirst { it.isDigit() }
        if (splitIndex <= 0 || splitIndex == text.length) return false

        val col = text.substring(0, splitIndex)
        val row = text.substring(splitIndex)
        if (col.length !in 1..3) return false
        if (!col.all { it.isLetter() }) return false
        if (!row.all { it.isDigit() }) return false

        return row.toIntOrNull()?.let { it >= 1 } == true
    }

    private companion object {
        private val R1C1_REGEX = Regex("^R[0-9]+C[0-9]+$")
    }
}
