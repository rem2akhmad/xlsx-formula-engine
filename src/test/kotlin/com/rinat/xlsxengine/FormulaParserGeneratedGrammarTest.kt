package com.rinat.xlsxengine

import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue
import kotlin.test.fail

class FormulaParserGeneratedGrammarTest {
    private val generator = FormulaGrammarStringGenerator()
    private val limit = 1000

    @Test
    fun `generator produces formulas of requested exact length`() {
        val exhaustiveLengths = listOf(1, 2, 3, 5, 10, 15, 20, 25)
        for (length in exhaustiveLengths) {
            val formulas = generator.generateExactLength(length, maxResults = limit)
                .ifEmpty { generator.generateExactLength(length, maxResults = 1) }
            assertTrue(formulas.isNotEmpty(), "No formulas generated for length=$length")
            assertTrue(formulas.size <= limit, "Generator exceeded limit for length=$length")
            formulas.forEach { formula ->
                assertEquals(length, formula.length, "Unexpected length for '$formula'")
            }
        }

        val longLengthFormula = generator.generateExactLength(250, maxResults = 1)
        assertTrue(longLengthFormula.isNotEmpty(), "No formulas generated for length=250")
        longLengthFormula.forEach { formula ->
            assertEquals(250, formula.length, "Unexpected length for '$formula'")
        }
    }

    @Test
    fun `generator can produce length 1000 formulas`() {
        val formulas = generator.generateExactLength(1000, maxResults = 1)
        assertTrue(formulas.isNotEmpty(), "No formulas generated for length=1000")
        formulas.forEach { formula ->
            assertEquals(1000, formula.length, "Unexpected length for '$formula'")
        }
    }

    @Test
    fun `parser accepts generated formulas from grammar`() {
        var parsedCount = 0
        val formulas = linkedSetOf<String>()
        formulas += generator.generateUpToLength(40, maxResults = limit - 1)
        formulas += generator.generateExactLength(1000, maxResults = 1)
        assertTrue(formulas.size <= limit, "Generator exceeded global test limit")

        for (formula in formulas) {
            try {
                parseFormula(formula)
                parsedCount++
            } catch (t: Throwable) {
                fail("Generated formula failed to parse: '$formula': ${t.message}")
            }
        }

        assertTrue(parsedCount > 100, "Too few generated parser checks: $parsedCount")
    }

    @Test
    fun `generated corpus includes key syntax constructs`() {
        val corpus = generator.generateUpToLength(40, maxResults = limit)
        assertTrue(corpus.size > 100, "Generated corpus is too small for a random grammar walk")
        assertTrue(corpus.any { it.startsWith("=") }, "Missing formulas with '=' prefix")
        assertTrue(corpus.any { it.contains('%') }, "Missing postfix percent formulas")
        assertTrue(corpus.any { it.contains('&') }, "Missing concat formulas")
        assertTrue(corpus.any { it.contains("R1C1") || it.contains("A1") }, "Missing cell reference formulas")
        assertTrue(corpus.any { it.contains('(') && it.contains(')') }, "Missing grouped/function-style formulas")
    }
}
