package com.rinat.xlsxengine

import kotlin.math.min
import kotlin.random.Random

/**
 * Generates formula strings by expanding a spec-inspired grammar.
 * Rule selection is randomized, while all produced strings stay grammar-valid.
 */
internal class FormulaGrammarStringGenerator(
    private val random: Random = Random.Default
) {
    private val grammar: Grammar = Grammar.specInspired()
    private val minLens: Map<String, Int> = grammar.computeMinLengths()

    fun generateExactLength(length: Int, maxResults: Int = Int.MAX_VALUE): Set<String> {
        require(length >= 0) { "length must be >= 0" }
        require(maxResults > 0) { "maxResults must be > 0" }

        // For very long targets, use a compact guaranteed-valid construction.
        if (length >= 128 && maxResults == 1) {
            return setOf(buildUnaryChainFormula(length))
        }

        val startMin = minLens[grammar.start] ?: Int.MAX_VALUE
        if (startMin > length) return emptySet()

        val out = linkedSetOf<String>()
        val maxAttempts = when {
            maxResults == 1 -> 64
            maxResults <= 32 -> maxResults * 80
            else -> maxResults * 120
        }

        repeat(maxAttempts) {
            if (out.size >= maxResults) return@repeat
            val candidate = randomDerivation(length)
            if (candidate != null) {
                out += candidate
            }
        }

        // Always return at least one valid string for feasible lengths.
        if (out.isEmpty() && length >= 1 && maxResults >= 1) {
            out += buildUnaryChainFormula(length)
        }

        return out
    }

    fun generateUpToLength(maxLength: Int, maxResults: Int = Int.MAX_VALUE): Set<String> {
        require(maxLength >= 0) { "maxLength must be >= 0" }
        require(maxResults > 0) { "maxResults must be > 0" }

        val out = linkedSetOf<String>()
        for (len in 0..maxLength) {
            if (out.size >= maxResults) break
            val remaining = maxResults - out.size
            out += generateExactLength(len, maxResults = remaining)
        }
        return out
    }

    private fun randomDerivation(targetLength: Int): String? {
        val state = mutableListOf<Symbol>(Symbol.NonTerminal(grammar.start))
        var guard = 0
        val maxSteps = (targetLength + 1) * 20

        while (guard++ < maxSteps) {
            val ntPositions = state.withIndex().filter { it.value is Symbol.NonTerminal }
            if (ntPositions.isEmpty()) {
                val len = terminalsLength(state)
                if (len == targetLength) {
                    return state.joinToString(separator = "") { (it as Symbol.Terminal).text }
                }
                return null
            }

            val currentTerminalLen = terminalsLength(state)
            if (currentTerminalLen > targetLength) return null
            if (minLength(state, minLens) > targetLength) return null

            val ntPos = ntPositions[random.nextInt(ntPositions.size)].index
            val nt = (state[ntPos] as Symbol.NonTerminal).name
            val productions = grammar.rules[nt].orEmpty()
            if (productions.isEmpty()) return null

            val feasible = productions.filter { prod ->
                val next = replaceAt(state, ntPos, prod)
                terminalsLength(next) <= targetLength && minLength(next, minLens) <= targetLength
            }
            if (feasible.isEmpty()) return null

            val chosen = feasible[random.nextInt(feasible.size)]
            val nextState = replaceAt(state, ntPos, chosen)
            state.clear()
            state.addAll(nextState)
        }

        return null
    }

    private fun replaceAt(state: List<Symbol>, index: Int, replacement: List<Symbol>): List<Symbol> {
        val out = ArrayList<Symbol>(state.size - 1 + replacement.size)
        out.addAll(state.subList(0, index))
        out.addAll(replacement)
        out.addAll(state.subList(index + 1, state.size))
        return out
    }

    private fun minLength(state: List<Symbol>, minLens: Map<String, Int>): Int {
        var sum = 0
        for (symbol in state) {
            when (symbol) {
                is Symbol.Terminal -> sum += symbol.text.length
                is Symbol.NonTerminal -> {
                    val v = minLens[symbol.name] ?: Int.MAX_VALUE
                    if (v == Int.MAX_VALUE) return Int.MAX_VALUE
                    sum += v
                }
            }
            if (sum < 0) return Int.MAX_VALUE
        }
        return sum
    }

    private fun terminalsLength(state: List<Symbol>): Int {
        var sum = 0
        for (symbol in state) {
            if (symbol is Symbol.Terminal) sum += symbol.text.length
        }
        return sum
    }

    private fun buildUnaryChainFormula(length: Int): String {
        require(length >= 1)
        if (length == 1) return "0"
        return buildString(length) {
            repeat(length - 1) {
                append(if (random.nextBoolean()) '+' else '-')
            }
            append('0')
        }
    }
}

private sealed interface Symbol {
    data class Terminal(val text: String) : Symbol
    data class NonTerminal(val name: String) : Symbol
}

private data class Grammar(
    val start: String,
    val rules: Map<String, List<List<Symbol>>>
) {
    fun computeMinLengths(): Map<String, Int> {
        val nonTerminals = rules.keys
        val minLens = nonTerminals.associateWith { Int.MAX_VALUE }.toMutableMap()

        var changed = true
        var guard = 0
        while (changed && guard < 10_000) {
            changed = false
            guard++

            for (nt in nonTerminals) {
                val current = minLens.getValue(nt)
                val best = rules.getValue(nt).minOf { production ->
                    var sum = 0
                    var impossible = false
                    for (symbol in production) {
                        when (symbol) {
                            is Symbol.Terminal -> sum += symbol.text.length
                            is Symbol.NonTerminal -> {
                                val v = minLens.getValue(symbol.name)
                                if (v == Int.MAX_VALUE) {
                                    impossible = true
                                    break
                                }
                                sum += v
                            }
                        }
                    }
                    if (impossible) Int.MAX_VALUE else sum
                }

                val next = min(current, best)
                if (next != current) {
                    minLens[nt] = next
                    changed = true
                }
            }
        }

        return minLens
    }

    companion object {
        fun specInspired(): Grammar {
            fun t(s: String) = Symbol.Terminal(s)
            fun n(s: String) = Symbol.NonTerminal(s)

            return Grammar(
                start = "Formula",
                rules = mapOf(
                    "Formula" to listOf(
                        listOf(n("Expr")),
                        listOf(t("="), n("Expr"))
                    ),
                    "Expr" to listOf(listOf(n("Comparison"))),
                    "Comparison" to listOf(listOf(n("Concat"), n("ComparisonTail"))),
                    "ComparisonTail" to listOf(
                        emptyList(),
                        listOf(t(">"), n("Concat"), n("ComparisonTail")),
                        listOf(t("<"), n("Concat"), n("ComparisonTail")),
                        listOf(t(">="), n("Concat"), n("ComparisonTail")),
                        listOf(t("<="), n("Concat"), n("ComparisonTail")),
                        listOf(t("="), n("Concat"), n("ComparisonTail")),
                        listOf(t("<>"), n("Concat"), n("ComparisonTail"))
                    ),
                    "Concat" to listOf(listOf(n("Add"), n("ConcatTail"))),
                    "ConcatTail" to listOf(
                        emptyList(),
                        listOf(t("&"), n("Add"), n("ConcatTail"))
                    ),
                    "Add" to listOf(listOf(n("Mul"), n("AddTail"))),
                    "AddTail" to listOf(
                        emptyList(),
                        listOf(t("+"), n("Mul"), n("AddTail")),
                        listOf(t("-"), n("Mul"), n("AddTail"))
                    ),
                    "Mul" to listOf(listOf(n("Pow"), n("MulTail"))),
                    "MulTail" to listOf(
                        emptyList(),
                        listOf(t("*"), n("Pow"), n("MulTail")),
                        listOf(t("/"), n("Pow"), n("MulTail"))
                    ),
                    "Pow" to listOf(listOf(n("Unary"), n("PowTail"))),
                    "PowTail" to listOf(
                        emptyList(),
                        listOf(t("^"), n("Unary"), n("PowTail"))
                    ),
                    "Unary" to listOf(
                        listOf(n("Primary")),
                        listOf(t("+"), n("Unary")),
                        listOf(t("-"), n("Unary"))
                    ),
                    "Primary" to listOf(listOf(n("Atom"), n("PercentTail"))),
                    "PercentTail" to listOf(
                        emptyList(),
                        listOf(t("%"), n("PercentTail"))
                    ),
                    "Atom" to listOf(
                        listOf(n("Number")),
                        listOf(n("Boolean")),
                        listOf(n("Error")),
                        listOf(n("Cell")),
                        listOf(n("Range")),
                        listOf(n("SheetCell")),
                        listOf(n("SheetRange")),
                        listOf(n("Range3D")),
                        listOf(n("DefinedName")),
                        listOf(n("FunctionCall")),
                        listOf(n("String")),
                        listOf(t("("), n("Expr"), t(")"))
                    ),
                    "FunctionCall" to listOf(
                        listOf(n("FunctionName"), t("("), n("ArgList"), t(")"))
                    ),
                    "ArgList" to listOf(
                        emptyList(),
                        listOf(n("ArgItem"), n("ArgListTail"))
                    ),
                    "ArgListTail" to listOf(
                        emptyList(),
                        listOf(t(","), n("ArgItem"), n("ArgListTail"))
                    ),
                    "ArgItem" to listOf(
                        emptyList(),
                        listOf(n("Expr"))
                    ),
                    "FunctionName" to listOf(
                        listOf(t("SUM")),
                        listOf(t("IF")),
                        listOf(t("MAX"))
                    ),
                    "DefinedName" to listOf(
                        listOf(t("X")),
                        listOf(t("RATE"))
                    ),
                    "Number" to listOf(
                        listOf(t("0")),
                        listOf(t("1")),
                        listOf(t("2.5"))
                    ),
                    "Boolean" to listOf(
                        listOf(t("TRUE")),
                        listOf(t("FALSE"))
                    ),
                    "Error" to listOf(
                        listOf(t("#REF!")),
                        listOf(t("#DIV/0!")),
                        listOf(t("#VALUE!")),
                        listOf(t("#NAME?"))
                    ),
                    "String" to listOf(
                        listOf(t("\"\"")),
                        listOf(t("\"A\""))
                    ),
                    "Cell" to listOf(
                        listOf(t("A1")),
                        listOf(t("B2")),
                        listOf(t("R1C1"))
                    ),
                    "Range" to listOf(
                        listOf(t("A1:B2")),
                        listOf(t("R1C1:R2C2"))
                    ),
                    "SheetCell" to listOf(
                        listOf(t("S!A1")),
                        listOf(t("'S 1'!B2"))
                    ),
                    "SheetRange" to listOf(
                        listOf(t("S!A1:B2"))
                    ),
                    "Range3D" to listOf(
                        listOf(t("S:T!A1:B2"))
                    )
                )
            )
        }
    }
}
