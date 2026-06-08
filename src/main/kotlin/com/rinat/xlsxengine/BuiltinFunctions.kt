package com.rinat.xlsxengine

import java.text.DecimalFormat
import java.text.DecimalFormatSymbols
import java.time.DayOfWeek
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.LocalTime
import java.time.YearMonth
import java.time.temporal.ChronoUnit
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException
import java.util.Locale
import kotlin.math.PI
import kotlin.math.acos
import kotlin.math.acosh
import kotlin.math.abs
import kotlin.math.asin
import kotlin.math.asinh
import kotlin.math.atan2
import kotlin.math.atanh
import kotlin.math.ceil
import kotlin.math.cosh
import kotlin.math.exp
import kotlin.math.floor
import kotlin.math.ln
import kotlin.math.log10
import kotlin.math.max
import kotlin.math.pow
import kotlin.math.round
import kotlin.math.sin
import kotlin.math.sinh
import kotlin.math.sqrt
import kotlin.math.tan
import kotlin.math.tanh
import kotlin.math.truncate

internal class IfFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val condition = context.evalScalar(currentSheet, args[0])
        if (condition is CellValue.ErrorValue) return EvalValue.Scalar(condition)

        val conditionValue = context.asBoolean(condition)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val branch = if (conditionValue) args[1] else args.getOrNull(2)
        if (branch == null) return EvalValue.Scalar(CellValue.BooleanValue(false))

        return when (val value = context.evalNode(currentSheet, branch)) {
            is EvalValue.Scalar -> value
            is EvalValue.Range,
            is EvalValue.Range3D -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class IfErrorFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val first = when (val value = context.evalNode(currentSheet, args[0])) {
            is EvalValue.Scalar -> value.value
            is EvalValue.Range,
            is EvalValue.Range3D -> CellValue.ErrorValue(ErrorCodes.VALUE)
        }

        if (first is CellValue.ErrorValue) {
            return when (val second = context.evalNode(currentSheet, args[1])) {
                is EvalValue.Scalar -> second
                is EvalValue.Range,
                is EvalValue.Range3D -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            }
        }

        return EvalValue.Scalar(first)
    }
}

internal class IfNaFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val first = when (val value = context.evalNode(currentSheet, args[0])) {
            is EvalValue.Scalar -> value.value
            is EvalValue.Range, is EvalValue.Range3D -> CellValue.ErrorValue(ErrorCodes.VALUE)
        }
        if (first is CellValue.ErrorValue && first.message.equals(ErrorCodes.NA, ignoreCase = true)) {
            return when (val second = context.evalNode(currentSheet, args[1])) {
                is EvalValue.Scalar -> second
                is EvalValue.Range, is EvalValue.Range3D -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            }
        }
        return EvalValue.Scalar(first)
    }
}

internal class AndFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        var result = true
        val err = context.forEachArgValue(currentSheet, args) { value ->
            val b = context.asBoolean(value) ?: run {
                result = false
                return@forEachArgValue
            }
            result = result && b
        }
        if (err != null) return EvalValue.Scalar(err)
        return EvalValue.Scalar(CellValue.BooleanValue(result))
    }
}

internal class OrFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        var result = false
        val err = context.forEachArgValue(currentSheet, args) { value ->
            val b = context.asBoolean(value) ?: run {
                result = result || false
                return@forEachArgValue
            }
            result = result || b
        }
        if (err != null) return EvalValue.Scalar(err)
        return EvalValue.Scalar(CellValue.BooleanValue(result))
    }
}

internal class NotFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        if (v is CellValue.ErrorValue) return EvalValue.Scalar(v)
        val b = context.asBoolean(v) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.BooleanValue(!b))
    }
}

internal class IsErrorFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        return EvalValue.Scalar(CellValue.BooleanValue(v is CellValue.ErrorValue))
    }
}

internal class IsErrFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        val result = v is CellValue.ErrorValue && !v.message.equals(ErrorCodes.NA, ignoreCase = true)
        return EvalValue.Scalar(CellValue.BooleanValue(result))
    }
}

internal class IsNaFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        val result = v is CellValue.ErrorValue && v.message.equals(ErrorCodes.NA, ignoreCase = true)
        return EvalValue.Scalar(CellValue.BooleanValue(result))
    }
}

internal class IsNumberFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        return EvalValue.Scalar(CellValue.BooleanValue(v is CellValue.NumberValue))
    }
}

internal class IsBlankFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        return EvalValue.Scalar(CellValue.BooleanValue(v == CellValue.Blank))
    }
}

internal class IsLogicalFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        return EvalValue.Scalar(CellValue.BooleanValue(v is CellValue.BooleanValue))
    }
}

internal class IsNonTextFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        val isNonText = v !is CellValue.TextValue
        return EvalValue.Scalar(CellValue.BooleanValue(isNonText))
    }
}

internal class IsTextFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val v = context.evalScalar(currentSheet, args[0])
        return EvalValue.Scalar(CellValue.BooleanValue(v is CellValue.TextValue))
    }
}

internal class IsEvenFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val n = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val k = floor(abs(n)).toLong()
        return EvalValue.Scalar(CellValue.BooleanValue(k % 2L == 0L))
    }
}

internal class IsOddFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val n = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val k = floor(abs(n)).toLong()
        return EvalValue.Scalar(CellValue.BooleanValue(k % 2L != 0L))
    }
}

internal class AreasFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        var count = 0
        for (arg in args) {
            when (arg) {
                is FormulaNode.CellRef, is FormulaNode.RangeRef -> count += 1
                is FormulaNode.Range3DRef -> count += 1
                else -> {
                    when (val ev = context.evalNode(currentSheet, arg)) {
                        is EvalValue.Range -> count += 1
                        is EvalValue.Range3D -> count += ev.ranges.size
                        is EvalValue.Scalar -> return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
                    }
                }
            }
        }
        return EvalValue.Scalar(CellValue.NumberValue(count.toDouble()))
    }
}

internal class TrueFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.BooleanValue(true))
    }
}

internal class FalseFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.BooleanValue(false))
    }
}

internal class ChooseFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size < 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val idx = context.evalScalar(currentSheet, args[0]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (idx < 1 || idx >= args.size) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return when (val pick = context.evalNode(currentSheet, args[idx])) {
            is EvalValue.Scalar -> pick
            is EvalValue.Range -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            is EvalValue.Range3D -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class MatchFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val lookup = context.evalScalar(currentSheet, args[0])
        if (lookup is CellValue.ErrorValue) return EvalValue.Scalar(lookup)
        val matchType = if (args.size == 3) {
            context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else 1

        val values = flattenSingleDimension(context, currentSheet, args[1])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))

        val idx = when (matchType) {
            0 -> values.indexOfFirst { compareForMatch(lookup, it) == 0 }.let { if (it >= 0) it + 1 else -1 }
            1 -> {
                var best = -1
                for (i in values.indices) {
                    if (compareForMatch(values[i], lookup) <= 0) best = i + 1
                }
                best
            }
            -1 -> {
                var best = -1
                for (i in values.indices) {
                    if (compareForMatch(values[i], lookup) >= 0) {
                        best = i + 1
                        break
                    }
                }
                best
            }
            else -> return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        }
        if (idx < 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        return EvalValue.Scalar(CellValue.NumberValue(idx.toDouble()))
    }
}

internal class VLookupFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..4) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val lookup = context.evalScalar(currentSheet, args[0])
        if (lookup is CellValue.ErrorValue) return EvalValue.Scalar(lookup)
        val table = context.evalNode(currentSheet, args[1]) as? EvalValue.Range
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        val colIndex = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (colIndex < 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val exact = if (args.size == 4) {
            val b = context.asBoolean(context.evalScalar(currentSheet, args[3]))
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            !b
        } else {
            false
        }

        val width = table.end.column - table.start.column + 1
        val height = table.end.row - table.start.row + 1
        if (colIndex > width) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))

        var targetRow = -1
        if (exact) {
            for (r in 0 until height) {
                val first = context.evaluateCell(table.sheet, CellAddress(table.start.row + r, table.start.column))
                if (compareForMatch(lookup, first) == 0) {
                    targetRow = r
                    break
                }
            }
        } else {
            for (r in 0 until height) {
                val first = context.evaluateCell(table.sheet, CellAddress(table.start.row + r, table.start.column))
                if (compareForMatch(first, lookup) <= 0) targetRow = r
            }
        }
        if (targetRow < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))

        val result = context.evaluateCell(
            table.sheet,
            CellAddress(table.start.row + targetRow, table.start.column + (colIndex - 1))
        )
        return EvalValue.Scalar(result)
    }
}

internal class HLookupFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..4) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val lookup = context.evalScalar(currentSheet, args[0])
        if (lookup is CellValue.ErrorValue) return EvalValue.Scalar(lookup)
        val table = context.evalNode(currentSheet, args[1]) as? EvalValue.Range
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        val rowIndex = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (rowIndex < 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val exact = if (args.size == 4) {
            val b = context.asBoolean(context.evalScalar(currentSheet, args[3]))
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            !b
        } else {
            false
        }

        val width = table.end.column - table.start.column + 1
        val height = table.end.row - table.start.row + 1
        if (rowIndex > height) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))

        var targetCol = -1
        if (exact) {
            for (c in 0 until width) {
                val first = context.evaluateCell(table.sheet, CellAddress(table.start.row, table.start.column + c))
                if (compareForMatch(lookup, first) == 0) {
                    targetCol = c
                    break
                }
            }
        } else {
            for (c in 0 until width) {
                val first = context.evaluateCell(table.sheet, CellAddress(table.start.row, table.start.column + c))
                if (compareForMatch(first, lookup) <= 0) targetCol = c
            }
        }
        if (targetCol < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))

        val result = context.evaluateCell(
            table.sheet,
            CellAddress(table.start.row + (rowIndex - 1), table.start.column + targetCol)
        )
        return EvalValue.Scalar(result)
    }
}

internal class LookupFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val lookupValue = context.evalScalar(currentSheet, args[0])
        if (lookupValue is CellValue.ErrorValue) return EvalValue.Scalar(lookupValue)

        val lookupVector = flattenSingleDimension(context, currentSheet, args[1])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        if (lookupVector.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))

        val resultVector = if (args.size == 3) {
            flattenSingleDimension(context, currentSheet, args[2])
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        } else {
            lookupVector
        }

        if (resultVector.size != lookupVector.size) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        }

        var idx = -1
        for (i in lookupVector.indices) {
            if (compareForMatch(lookupVector[i], lookupValue) <= 0) idx = i
        }
        if (idx < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        return EvalValue.Scalar(resultVector[idx])
    }
}

internal class OffsetFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..5) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val base = resolveRangeReference(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        val rowOffset = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val colOffset = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val baseHeight = base.end.row - base.start.row + 1
        val baseWidth = base.end.column - base.start.column + 1
        val height = if (args.size >= 4) {
            context.evalScalar(currentSheet, args[3]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else baseHeight
        val width = if (args.size >= 5) {
            context.evalScalar(currentSheet, args[4]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else baseWidth

        if (height <= 0 || width <= 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        val start = CellAddress(base.start.row + rowOffset, base.start.column + colOffset)
        val end = CellAddress(start.row + height - 1, start.column + width - 1)
        return EvalValue.Range(base.sheet, start, end)
    }
}

internal class IndirectFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val textValue = context.evalScalar(currentSheet, args[0])
        if (textValue is CellValue.ErrorValue) return EvalValue.Scalar(textValue)
        val a1Style = if (args.size == 2) {
            context.asBoolean(context.evalScalar(currentSheet, args[1]))
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else true
        if (!a1Style) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val text = toTextValue(textValue).trim()
        return try {
            context.evalNode(currentSheet, parseFormula(text))
        } catch (_: Throwable) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        }
    }
}

internal class SumIfFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val range = context.evalNode(currentSheet, args[0]) as? EvalValue.Range
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = context.evalScalar(currentSheet, args[1])
        if (criteria is CellValue.ErrorValue) return EvalValue.Scalar(criteria)
        val sumRange = if (args.size == 3) {
            context.evalNode(currentSheet, args[2]) as? EvalValue.Range ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            range
        }

        val predicate = criteriaPredicate(criteria)
        var sum = 0.0
        val rows = range.end.row - range.start.row + 1
        val cols = range.end.column - range.start.column + 1
        for (r in 0 until rows) {
            for (c in 0 until cols) {
                val cv = context.evaluateCell(range.sheet, CellAddress(range.start.row + r, range.start.column + c))
                if (predicate(cv)) {
                    val sv = context.evaluateCell(sumRange.sheet, CellAddress(sumRange.start.row + r, sumRange.start.column + c))
                    sv.asNumberOrNull()?.let { sum += it }
                }
            }
        }
        return EvalValue.Scalar(CellValue.NumberValue(sum))
    }
}

internal class CountIfFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val range = context.evalNode(currentSheet, args[0]) as? EvalValue.Range
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = context.evalScalar(currentSheet, args[1])
        if (criteria is CellValue.ErrorValue) return EvalValue.Scalar(criteria)
        val predicate = criteriaPredicate(criteria)
        var count = 0.0
        for (v in context.iterateRange(range.sheet, range.start, range.end)) {
            if (predicate(v)) count += 1.0
        }
        return EvalValue.Scalar(CellValue.NumberValue(count))
    }
}

internal class AverageIfFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val range = context.evalNode(currentSheet, args[0]) as? EvalValue.Range
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = context.evalScalar(currentSheet, args[1])
        if (criteria is CellValue.ErrorValue) return EvalValue.Scalar(criteria)
        val avgRange = if (args.size == 3) {
            context.evalNode(currentSheet, args[2]) as? EvalValue.Range ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            range
        }
        val predicate = criteriaPredicate(criteria)
        var sum = 0.0
        var count = 0
        val rows = range.end.row - range.start.row + 1
        val cols = range.end.column - range.start.column + 1
        for (r in 0 until rows) {
            for (c in 0 until cols) {
                val cv = context.evaluateCell(range.sheet, CellAddress(range.start.row + r, range.start.column + c))
                if (predicate(cv)) {
                    val sv = context.evaluateCell(avgRange.sheet, CellAddress(avgRange.start.row + r, avgRange.start.column + c))
                    val n = sv.asNumberOrNull()
                    if (n != null) {
                        sum += n
                        count++
                    }
                }
            }
        }
        if (count == 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        return EvalValue.Scalar(CellValue.NumberValue(sum / count))
    }
}

internal class SumIfsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size < 3 || args.size % 2 == 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val sumRange = resolveRangeReference(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = parseCriteriaPairs(context, currentSheet, args.drop(1)) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (!criteria.all { sameRangeShape(sumRange, it.range) }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val rows = sumRange.end.row - sumRange.start.row + 1
        val cols = sumRange.end.column - sumRange.start.column + 1
        var sum = 0.0
        for (r in 0 until rows) {
            for (c in 0 until cols) {
                if (criteria.all { pair ->
                        val v = context.evaluateCell(pair.range.sheet, CellAddress(pair.range.start.row + r, pair.range.start.column + c))
                        pair.predicate(v)
                    }) {
                    val sv = context.evaluateCell(sumRange.sheet, CellAddress(sumRange.start.row + r, sumRange.start.column + c))
                    sv.asNumberOrNull()?.let { sum += it }
                }
            }
        }
        return EvalValue.Scalar(CellValue.NumberValue(sum))
    }
}

internal class CountIfsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size < 2 || args.size % 2 != 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = parseCriteriaPairs(context, currentSheet, args) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val base = criteria.first().range
        if (!criteria.all { sameRangeShape(base, it.range) }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val rows = base.end.row - base.start.row + 1
        val cols = base.end.column - base.start.column + 1
        var count = 0.0
        for (r in 0 until rows) {
            for (c in 0 until cols) {
                val ok = criteria.all { pair ->
                    val v = context.evaluateCell(pair.range.sheet, CellAddress(pair.range.start.row + r, pair.range.start.column + c))
                    pair.predicate(v)
                }
                if (ok) count += 1.0
            }
        }
        return EvalValue.Scalar(CellValue.NumberValue(count))
    }
}

internal class AverageIfsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size < 3 || args.size % 2 == 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val avgRange = resolveRangeReference(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val criteria = parseCriteriaPairs(context, currentSheet, args.drop(1)) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (!criteria.all { sameRangeShape(avgRange, it.range) }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val rows = avgRange.end.row - avgRange.start.row + 1
        val cols = avgRange.end.column - avgRange.start.column + 1
        var sum = 0.0
        var count = 0
        for (r in 0 until rows) {
            for (c in 0 until cols) {
                if (criteria.all { pair ->
                        val v = context.evaluateCell(pair.range.sheet, CellAddress(pair.range.start.row + r, pair.range.start.column + c))
                        pair.predicate(v)
                    }) {
                    val sv = context.evaluateCell(avgRange.sheet, CellAddress(avgRange.start.row + r, avgRange.start.column + c))
                    val n = sv.asNumberOrNull()
                    if (n != null) {
                        sum += n
                        count++
                    }
                }
            }
        }
        if (count == 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        return EvalValue.Scalar(CellValue.NumberValue(sum / count))
    }
}

internal class IndexFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val source = context.evalNode(currentSheet, args[0])
        val range = when (source) {
            is EvalValue.Range -> source
            is EvalValue.Range3D -> source.ranges.firstOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
            is EvalValue.Scalar -> return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }

        val rows = range.end.row - range.start.row + 1
        val cols = range.end.column - range.start.column + 1

        val rowNum = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val colNum = if (args.size == 3) {
            context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            if (rows == 1) rowNum else 1
        }

        val effectiveRow = if (rows == 1 && args.size == 2) 1 else rowNum

        if (effectiveRow !in 1..rows || colNum !in 1..cols) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.REF))
        }

        val targetAddress = CellAddress(
            row = range.start.row + (effectiveRow - 1),
            column = range.start.column + (colNum - 1)
        )
        return EvalValue.Scalar(context.evaluateCell(range.sheet, targetAddress))
    }
}

internal class SumFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var sum = 0.0
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                directNumeric(value)?.let { sum += it }
            },
            onRange = { value ->
                rangeNumeric(value)?.let { sum += it }
            }
        )
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(sum))
    }
}

internal class AverageFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var sum = 0.0
        var count = 0
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                directNumeric(value)?.let {
                    sum += it
                    count++
                }
            },
            onRange = { value ->
                rangeNumeric(value)?.let {
                    sum += it
                    count++
                }
            }
        )

        return when {
            error != null -> EvalValue.Scalar(error)
            count == 0 -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
            else -> EvalValue.Scalar(CellValue.NumberValue(sum / count))
        }
    }
}

internal class MinFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var min: Double? = null
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                directNumeric(value)?.let { candidate ->
                    min = if (min == null) candidate else kotlin.math.min(min!!, candidate)
                }
            },
            onRange = { value ->
                rangeNumeric(value)?.let { candidate ->
                    min = if (min == null) candidate else kotlin.math.min(min!!, candidate)
                }
            }
        )

        return when {
            error != null -> EvalValue.Scalar(error)
            min == null -> EvalValue.Scalar(CellValue.NumberValue(0.0))
            else -> EvalValue.Scalar(CellValue.NumberValue(min!!))
        }
    }
}

internal class MaxFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var max: Double? = null
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                directNumeric(value)?.let { candidate ->
                    max = if (max == null) candidate else kotlin.math.max(max!!, candidate)
                }
            },
            onRange = { value ->
                rangeNumeric(value)?.let { candidate ->
                    max = if (max == null) candidate else kotlin.math.max(max!!, candidate)
                }
            }
        )

        return when {
            error != null -> EvalValue.Scalar(error)
            max == null -> EvalValue.Scalar(CellValue.NumberValue(0.0))
            else -> EvalValue.Scalar(CellValue.NumberValue(max!!))
        }
    }
}

internal class MaxaFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var max: Double? = null
        val error = context.forEachArgValue(currentSheet, args) { value ->
            val candidate = when (value) {
                is CellValue.NumberValue -> value.value
                is CellValue.BooleanValue -> if (value.value) 1.0 else 0.0
                is CellValue.TextValue -> 0.0
                CellValue.Blank -> null
                is CellValue.ErrorValue -> null
            }
            if (candidate != null) {
                max = if (max == null) candidate else kotlin.math.max(max!!, candidate)
            }
        }
        return when {
            error != null -> EvalValue.Scalar(error)
            max == null -> EvalValue.Scalar(CellValue.NumberValue(0.0))
            else -> EvalValue.Scalar(CellValue.NumberValue(max!!))
        }
    }
}

internal class AveDevFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = mutableListOf<Double>()
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { values += it } },
            onRange = { value -> rangeNumeric(value)?.let { values += it } }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        val mean = values.average()
        val avgDev = values.sumOf { abs(it - mean) } / values.size
        return EvalValue.Scalar(CellValue.NumberValue(avgDev))
    }
}

internal class CorrelFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val xRaw = flattenSingleDimension(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val yRaw = flattenSingleDimension(context, currentSheet, args[1])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val x = xRaw.mapNotNull { rangeNumeric(it) }
        val y = yRaw.mapNotNull { rangeNumeric(it) }
        if (x.size != y.size) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        if (x.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        val xMean = x.average()
        val yMean = y.average()
        var numerator = 0.0
        var xSq = 0.0
        var ySq = 0.0
        for (i in x.indices) {
            val dx = x[i] - xMean
            val dy = y[i] - yMean
            numerator += dx * dy
            xSq += dx * dx
            ySq += dy * dy
        }
        if (xSq == 0.0 || ySq == 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        return EvalValue.Scalar(CellValue.NumberValue(numerator / sqrt(xSq * ySq)))
    }
}

internal class GeomeanFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = mutableListOf<Double>()
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { values += it } },
            onRange = { value -> rangeNumeric(value)?.let { values += it } }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        if (values.any { it <= 0.0 }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val mean = exp(values.sumOf { ln(it) } / values.size)
        return EvalValue.Scalar(CellValue.NumberValue(mean))
    }
}

internal class HarmeanFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = mutableListOf<Double>()
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { values += it } },
            onRange = { value -> rangeNumeric(value)?.let { values += it } }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        if (values.any { it <= 0.0 }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val sumReciprocal = values.sumOf { 1.0 / it }
        return EvalValue.Scalar(CellValue.NumberValue(values.size / sumReciprocal))
    }
}

internal class ModeFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = mutableListOf<Double>()
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { values += it } },
            onRange = { value -> rangeNumeric(value)?.let { values += it } }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        val freq = linkedMapOf<Double, Int>()
        for (v in values) freq[v] = (freq[v] ?: 0) + 1
        val best = freq.entries.maxWithOrNull(compareBy<Map.Entry<Double, Int>> { it.value }.thenBy { -it.key })
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        if (best.value < 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        return EvalValue.Scalar(CellValue.NumberValue(best.key))
    }
}

internal class CountFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var count = 0
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                if (directNumeric(value) != null) count++
            },
            onRange = { value ->
                if (rangeNumeric(value) != null) count++
            }
        )
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(count.toDouble()))
    }
}

internal class CountaFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var count = 0
        val error = context.forEachArgValue(currentSheet, args) { value ->
            if (value != CellValue.Blank) count++
        }
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(count.toDouble()))
    }
}

internal class CountBlankFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var count = 0
        val error = context.forEachArgValue(currentSheet, args) { value ->
            if (value == CellValue.Blank || (value is CellValue.TextValue && value.value.isEmpty())) count++
        }
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(count.toDouble()))
    }
}

internal class AbsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = values.first().asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(abs(number)))
    }
}

internal class AcosFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number < -1.0 || number > 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(acos(number)))
    }
}

internal class AcoshFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number < 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(acosh(number)))
    }
}

internal class AsinFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number < -1.0 || number > 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(asin(number)))
    }
}

internal class AsinhFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(asinh(number)))
    }
}

internal class Atan2Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val x = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val y = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (x == 0.0 && y == 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        return EvalValue.Scalar(CellValue.NumberValue(atan2(y, x)))
    }
}

internal class AtanhFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number <= -1.0 || number >= 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(atanh(number)))
    }
}

internal class CoshFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(cosh(number)))
    }
}

internal class RadiansFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val angle = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(angle * PI / 180.0))
    }
}

internal class SinFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(sin(number)))
    }
}

internal class SinhFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(sinh(number)))
    }
}

internal class TanFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(tan(number)))
    }
}

internal class TanhFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(tanh(number)))
    }
}

internal class SqrtPiFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(sqrt(number * PI)))
    }
}

internal class TruncFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val digits = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            0
        }
        val factor = 10.0.pow(digits)
        val truncated = truncate(number * factor) / factor
        return EvalValue.Scalar(CellValue.NumberValue(truncated))
    }
}

internal class QuotientFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val numerator = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val denominator = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (denominator == 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        return EvalValue.Scalar(CellValue.NumberValue(truncate(numerator / denominator)))
    }
}

internal class ColumnFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val bounds = resolveReferenceBounds(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(bounds.start.column.toDouble()))
    }
}

internal class ColumnsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val bounds = resolveReferenceBounds(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val width = bounds.end.column - bounds.start.column + 1
        return EvalValue.Scalar(CellValue.NumberValue(width.toDouble()))
    }
}

internal class RowsFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val bounds = resolveReferenceBounds(context, currentSheet, args[0])
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val height = bounds.end.row - bounds.start.row + 1
        return EvalValue.Scalar(CellValue.NumberValue(height.toDouble()))
    }
}

internal class LnFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val number = values.first().asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number <= 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(kotlin.math.ln(number)))
    }
}

internal class SqrtFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val number = values.first().asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(sqrt(number)))
    }
}

internal class ExpFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val number = values.first().asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(exp(number)))
    }
}

internal class NormsDistFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val z = values.first().asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val cdf = normalCdf(z)
        return EvalValue.Scalar(CellValue.NumberValue(cdf))
    }

    // Abramowitz-Stegun style approximation for standard normal CDF.
    private fun normalCdf(x: Double): Double {
        val sign = if (x < 0.0) -1.0 else 1.0
        val absX = kotlin.math.abs(x) / kotlin.math.sqrt(2.0)
        val t = 1.0 / (1.0 + 0.3275911 * absX)
        val a1 = 0.254829592
        val a2 = -0.284496736
        val a3 = 1.421413741
        val a4 = -1.453152027
        val a5 = 1.061405429
        val erfApprox = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * kotlin.math.exp(-(absX * absX))
        val erf = sign * erfApprox
        return 0.5 * (1.0 + erf)
    }
}

internal class PowerFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val base = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val exponent = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(base.pow(exponent)))
    }
}

internal class IntFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(floor(value)))
    }
}

internal class ModFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val divisor = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (divisor == 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
        val result = number - divisor * floor(number / divisor)
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class RoundFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val digits = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            0
        }
        return EvalValue.Scalar(CellValue.NumberValue(roundHalfAwayFromZero(number, digits)))
    }
}

internal class RoundUpFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val digits = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            0
        }
        return EvalValue.Scalar(CellValue.NumberValue(roundAwayFromZero(number, digits)))
    }
}

internal class RoundDownFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val digits = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            0
        }
        return EvalValue.Scalar(CellValue.NumberValue(roundTowardZero(number, digits)))
    }
}

internal class LenFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        return EvalValue.Scalar(CellValue.NumberValue(toTextValue(value).length.toDouble()))
    }
}

internal class LeftFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val text = context.evalScalar(currentSheet, args[0])
        if (text is CellValue.ErrorValue) return EvalValue.Scalar(text)
        val count = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else 1
        if (count < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val s = toTextValue(text)
        return EvalValue.Scalar(CellValue.TextValue(s.take(count)))
    }
}

internal class LeftBFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        // DBCS byte-aware behavior is reduced to char-based behavior in this engine.
        return LeftFunction().evaluate(context, currentSheet, args)
    }
}

internal class RightFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val text = context.evalScalar(currentSheet, args[0])
        if (text is CellValue.ErrorValue) return EvalValue.Scalar(text)
        val count = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else 1
        if (count < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val s = toTextValue(text)
        return EvalValue.Scalar(CellValue.TextValue(s.takeLast(count)))
    }
}

internal class MidFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val text = context.evalScalar(currentSheet, args[0])
        if (text is CellValue.ErrorValue) return EvalValue.Scalar(text)
        val start = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val count = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (start < 1 || count < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val s = toTextValue(text)
        val from = minOf(start - 1, s.length)
        val to = minOf(from + count, s.length)
        return EvalValue.Scalar(CellValue.TextValue(s.substring(from, to)))
    }
}

internal class MidBFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        // DBCS byte-aware behavior is reduced to char-based behavior in this engine.
        return MidFunction().evaluate(context, currentSheet, args)
    }
}

internal class LenBFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        // DBCS byte-aware behavior is reduced to char-based behavior in this engine.
        return LenFunction().evaluate(context, currentSheet, args)
    }
}

internal class ConcatenateFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val out = StringBuilder()
        val err = context.forEachArgValue(currentSheet, args) { value ->
            out.append(toTextValue(value))
        }
        if (err != null) return EvalValue.Scalar(err)
        return EvalValue.Scalar(CellValue.TextValue(out.toString()))
    }
}

internal class UpperFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        return EvalValue.Scalar(CellValue.TextValue(toTextValue(value).uppercase()))
    }
}

internal class LowerFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        return EvalValue.Scalar(CellValue.TextValue(toTextValue(value).lowercase()))
    }
}

internal class TrimFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val normalized = toTextValue(value).trim().split(Regex(" +")).filter { it.isNotEmpty() }.joinToString(" ")
        return EvalValue.Scalar(CellValue.TextValue(normalized))
    }
}

internal class ReptFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val text = context.evalScalar(currentSheet, args[0])
        if (text is CellValue.ErrorValue) return EvalValue.Scalar(text)
        val count = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (count < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.TextValue(toTextValue(text).repeat(count)))
    }
}

internal class ExactFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val left = context.evalScalar(currentSheet, args[0])
        val right = context.evalScalar(currentSheet, args[1])
        if (left is CellValue.ErrorValue) return EvalValue.Scalar(left)
        if (right is CellValue.ErrorValue) return EvalValue.Scalar(right)
        return EvalValue.Scalar(CellValue.BooleanValue(toTextValue(left) == toTextValue(right)))
    }
}

internal class ValueFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val num = when (value) {
            is CellValue.NumberValue -> value.value
            is CellValue.BooleanValue -> if (value.value) 1.0 else 0.0
            is CellValue.TextValue -> value.value.trim().toDoubleOrNull()
            CellValue.Blank -> 0.0
            is CellValue.ErrorValue -> null
        } ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(num))
    }
}

internal class CharFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (number !in 1..255) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.TextValue(number.toChar().toString()))
    }
}

internal class CleanFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val cleaned = toTextValue(value).filter { it.code >= 32 }
        return EvalValue.Scalar(CellValue.TextValue(cleaned))
    }
}

internal class CodeFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val text = toTextValue(value)
        if (text.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(text.first().code.toDouble()))
    }
}

internal class ProperFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val input = toTextValue(value)
        val sb = StringBuilder(input.length)
        var capitalize = true
        for (ch in input) {
            if (ch.isLetter()) {
                sb.append(if (capitalize) ch.uppercaseChar() else ch.lowercaseChar())
                capitalize = false
            } else {
                sb.append(ch)
                capitalize = true
            }
        }
        return EvalValue.Scalar(CellValue.TextValue(sb.toString()))
    }
}

internal class ReplaceFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 4) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val oldTextValue = context.evalScalar(currentSheet, args[0])
        if (oldTextValue is CellValue.ErrorValue) return EvalValue.Scalar(oldTextValue)
        val oldText = toTextValue(oldTextValue)
        val start = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val numChars = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val newTextValue = context.evalScalar(currentSheet, args[3])
        if (newTextValue is CellValue.ErrorValue) return EvalValue.Scalar(newTextValue)
        val newText = toTextValue(newTextValue)
        if (start < 1 || numChars < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val from = minOf(start - 1, oldText.length)
        val to = minOf(from + numChars, oldText.length)
        val result = oldText.substring(0, from) + newText + oldText.substring(to)
        return EvalValue.Scalar(CellValue.TextValue(result))
    }
}

internal class TextFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val value = context.evalScalar(currentSheet, args[0])
        if (value is CellValue.ErrorValue) return EvalValue.Scalar(value)
        val formatValue = context.evalScalar(currentSheet, args[1])
        if (formatValue is CellValue.ErrorValue) return EvalValue.Scalar(formatValue)
        val format = toTextValue(formatValue)
        val number = value.asNumberOrNull()
        if (number == null) return EvalValue.Scalar(CellValue.TextValue(toTextValue(value)))
        return EvalValue.Scalar(CellValue.TextValue(applyTextFormat(number, format)))
    }
}

internal class DollarFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val decimals = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            2
        }
        val rounded = roundHalfAwayFromZero(number, decimals)
        val text = formatNumber(rounded, decimals, useGrouping = true)
        return EvalValue.Scalar(CellValue.TextValue("$$text"))
    }
}

internal class FixedFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val decimals = if (args.size >= 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            2
        }
        val noCommas = if (args.size >= 3) {
            context.asBoolean(context.evalScalar(currentSheet, args[2]))
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            false
        }
        val rounded = roundHalfAwayFromZero(number, decimals)
        val text = formatNumber(rounded, decimals, useGrouping = !noCommas)
        return EvalValue.Scalar(CellValue.TextValue(text))
    }
}

internal class FindFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return findOrSearch(context, currentSheet, args, caseSensitive = true)
    }
}

internal class SearchFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return findOrSearch(context, currentSheet, args, caseSensitive = false)
    }
}

internal class SubstituteFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..4) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val textV = context.evalScalar(currentSheet, args[0])
        val oldV = context.evalScalar(currentSheet, args[1])
        val newV = context.evalScalar(currentSheet, args[2])
        if (textV is CellValue.ErrorValue) return EvalValue.Scalar(textV)
        if (oldV is CellValue.ErrorValue) return EvalValue.Scalar(oldV)
        if (newV is CellValue.ErrorValue) return EvalValue.Scalar(newV)
        val text = toTextValue(textV)
        val old = toTextValue(oldV)
        val repl = toTextValue(newV)

        if (args.size == 3) {
            return EvalValue.Scalar(CellValue.TextValue(text.replace(old, repl)))
        }

        val instance = context.evalScalar(currentSheet, args[3]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (instance <= 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (old.isEmpty()) return EvalValue.Scalar(CellValue.TextValue(text))

        var seen = 0
        var idx = 0
        val out = StringBuilder()
        while (idx < text.length) {
            val at = text.indexOf(old, idx)
            if (at < 0) {
                out.append(text.substring(idx))
                break
            }
            out.append(text.substring(idx, at))
            seen++
            if (seen == instance) {
                out.append(repl)
                out.append(text.substring(at + old.length))
                return EvalValue.Scalar(CellValue.TextValue(out.toString()))
            } else {
                out.append(old)
                idx = at + old.length
            }
        }
        return EvalValue.Scalar(CellValue.TextValue(out.toString()))
    }
}

internal class TimeFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val hour = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val minute = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val second = context.evalScalar(currentSheet, args[2]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (hour < 0.0 || minute < 0.0 || second < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val totalSeconds = hour * 3600.0 + minute * 60.0 + second
        val normalized = ((totalSeconds % 86400.0) + 86400.0) % 86400.0
        return EvalValue.Scalar(CellValue.NumberValue(normalized / 86400.0))
    }
}

internal class TimeValueFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val input = context.evalScalar(currentSheet, args[0])
        if (input is CellValue.ErrorValue) return EvalValue.Scalar(input)
        val text = toTextValue(input).trim()
        val localTime = parseTimeText(text) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val seconds = localTime.toSecondOfDay().toDouble() + (localTime.nano / 1_000_000_000.0)
        return EvalValue.Scalar(CellValue.NumberValue(seconds / 86400.0))
    }
}

internal class DateFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val year = context.evalScalar(currentSheet, args[0]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val month = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val day = context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return try {
            val base = LocalDate.of(year, 1, 1).plusMonths((month - 1).toLong()).plusDays((day - 1).toLong())
            EvalValue.Scalar(CellValue.NumberValue(toExcelSerial(base)))
        } catch (_: Throwable) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }
    }
}

internal class DateDifFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val startSerial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val endSerial = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val unitValue = context.evalScalar(currentSheet, args[2])
        if (unitValue is CellValue.ErrorValue) return EvalValue.Scalar(unitValue)
        val unit = toTextValue(unitValue).uppercase(Locale.US)
        val start = fromExcelSerial(startSerial)
        val end = fromExcelSerial(endSerial)
        if (end.isBefore(start)) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))

        val result = when (unit) {
            "D" -> ChronoUnit.DAYS.between(start, end).toDouble()
            "M" -> fullMonthsBetween(start, end).toDouble()
            "Y" -> fullYearsBetween(start, end).toDouble()
            "MD" -> datedifMd(start, end).toDouble()
            "YM" -> (fullMonthsBetween(start, end) % 12).toDouble()
            "YD" -> ChronoUnit.DAYS.between(start.withYear(end.year), end).let { d ->
                if (d >= 0) d else ChronoUnit.DAYS.between(start.withYear(end.year - 1), end)
            }.toDouble()

            else -> return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class Days360Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val startSerial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val endSerial = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val methodUs = if (args.size == 3) {
            val european = context.asBoolean(context.evalScalar(currentSheet, args[2]))
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
            !european
        } else {
            true
        }
        val start = fromExcelSerial(startSerial)
        val end = fromExcelSerial(endSerial)
        val days = days360(start, end, methodUs)
        return EvalValue.Scalar(CellValue.NumberValue(days.toDouble()))
    }
}

internal class EDateFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val startSerial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val months = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return try {
            val start = fromExcelSerial(startSerial)
            val shifted = start.plusMonths(months.toLong())
            EvalValue.Scalar(CellValue.NumberValue(toExcelSerial(shifted)))
        } catch (_: Throwable) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }
    }
}

internal class EoMonthFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val startSerial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val months = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return try {
            val start = fromExcelSerial(startSerial)
            val ym = YearMonth.of(start.year, start.month).plusMonths(months.toLong())
            val end = ym.atEndOfMonth()
            EvalValue.Scalar(CellValue.NumberValue(toExcelSerial(end)))
        } catch (_: Throwable) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }
    }
}

internal class NetworkDaysFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val startSerial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val endSerial = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val holidays = if (args.size == 3) collectHolidayDates(context, currentSheet, args[2]) else emptySet()

        var start = fromExcelSerial(startSerial)
        var end = fromExcelSerial(endSerial)
        val step = if (start <= end) 1 else -1
        if (step < 0) {
            val tmp = start
            start = end
            end = tmp
        }

        var count = 0
        var d = start
        while (!d.isAfter(end)) {
            val weekend = d.dayOfWeek == DayOfWeek.SATURDAY || d.dayOfWeek == DayOfWeek.SUNDAY
            if (!weekend && d !in holidays) count++
            d = d.plusDays(1)
        }
        val result = if (step > 0) count else -count
        return EvalValue.Scalar(CellValue.NumberValue(result.toDouble()))
    }
}

internal class YearFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val date = extractDateArg(context, currentSheet, args) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(date.year.toDouble()))
    }
}

internal class MonthFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val date = extractDateArg(context, currentSheet, args) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(date.monthValue.toDouble()))
    }
}

internal class HourFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val fraction = extractTimeFraction(context, currentSheet, args)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val hours = floor(fraction * 24.0).toInt()
        return EvalValue.Scalar(CellValue.NumberValue(hours.toDouble()))
    }
}

internal class SecondFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val fraction = extractTimeFraction(context, currentSheet, args)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val totalSeconds = floor(fraction * 86400.0 + 1e-9).toInt()
        val seconds = totalSeconds % 60
        return EvalValue.Scalar(CellValue.NumberValue(seconds.toDouble()))
    }
}

internal class DayFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val date = extractDateArg(context, currentSheet, args) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(date.dayOfMonth.toDouble()))
    }
}

internal class TodayFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(toExcelSerial(LocalDate.now())))
    }
}

internal class NowFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val dt = LocalDateTime.now()
        val datePart = toExcelSerial(dt.toLocalDate())
        val timePart = dt.toLocalTime().toSecondOfDay() / 86400.0
        return EvalValue.Scalar(CellValue.NumberValue(datePart + timePart))
    }
}

internal class RandFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(kotlin.random.Random.nextDouble()))
    }
}

internal class CombinFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val n = context.evalScalar(currentSheet, args[0]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val k = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (n < 0 || k < 0 || k > n) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val r = minOf(k, n - k)
        var result = 1.0
        for (i in 1..r) {
            result = result * (n - r + i).toDouble() / i.toDouble()
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class GcdFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        var current = 0L
        var invalid = false
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                val n = directNumeric(value) ?: return@forEachArgumentValueWithSource
                if (n < 0.0) {
                    invalid = true
                    return@forEachArgumentValueWithSource
                }
                current = gcdLong(current, floor(n).toLong())
            },
            onRange = { value ->
                val n = rangeNumeric(value) ?: return@forEachArgumentValueWithSource
                if (n < 0.0) {
                    invalid = true
                    return@forEachArgumentValueWithSource
                }
                current = gcdLong(current, floor(n).toLong())
            }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (invalid) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(current.toDouble()))
    }
}

internal class GeStepFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty() || args.size > 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val step = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else 0.0
        return EvalValue.Scalar(CellValue.NumberValue(if (number >= step) 1.0 else 0.0))
    }
}

internal class SumProductFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val vectors = mutableListOf<List<Double>>()
        for (arg in args) {
            when (val evaluated = context.evalNode(currentSheet, arg)) {
                is EvalValue.Scalar -> {
                    val n = evaluated.value.asNumberOrNull() ?: 0.0
                    vectors += listOf(n)
                }

                is EvalValue.Range -> {
                    val values = context.iterateRange(evaluated.sheet, evaluated.start, evaluated.end)
                        .map { it.asNumberOrNull() ?: 0.0 }
                        .toList()
                    vectors += values
                }

                is EvalValue.Range3D -> {
                    val values = mutableListOf<Double>()
                    for (range in evaluated.ranges) {
                        values += context.iterateRange(range.sheet, range.start, range.end)
                            .map { it.asNumberOrNull() ?: 0.0 }
                            .toList()
                    }
                    vectors += values
                }
            }
        }
        val size = vectors.maxOf { it.size }
        if (vectors.any { it.size != 1 && it.size != size }) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        var sum = 0.0
        for (i in 0 until size) {
            var product = 1.0
            for (v in vectors) {
                product *= if (v.size == 1) v[0] else v[i]
            }
            sum += product
        }
        return EvalValue.Scalar(CellValue.NumberValue(sum))
    }
}

internal class SumX2My2Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return evalPairwiseArrayOp(context, currentSheet, args) { x, y -> x * x - y * y }
    }
}

internal class SumX2Py2Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return evalPairwiseArrayOp(context, currentSheet, args) { x, y -> x * x + y * y }
    }
}

internal class SumXMy2Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return evalPairwiseArrayOp(context, currentSheet, args) { x, y ->
            val d = x - y
            d * d
        }
    }
}

internal class RankFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val values = flattenSingleDimension(context, currentSheet, args[1])
            ?.mapNotNull { it.asNumberOrNull() }
            ?.takeIf { it.isNotEmpty() }
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
        val order = if (args.size == 3) {
            context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else 0
        val rank = if (order == 0) {
            1 + values.count { it > number }
        } else {
            1 + values.count { it < number }
        }
        return EvalValue.Scalar(CellValue.NumberValue(rank.toDouble()))
    }
}

internal class PercentileFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val values = flattenSingleDimension(context, currentSheet, args[0])
            ?.mapNotNull { it.asNumberOrNull() }
            ?.sorted()
            ?.takeIf { it.isNotEmpty() }
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val k = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (k < 0.0 || k > 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        if (values.size == 1) return EvalValue.Scalar(CellValue.NumberValue(values[0]))
        val n = values.size
        val idx = (n - 1) * k
        val lo = floor(idx).toInt()
        val hi = ceil(idx).toInt()
        val result = if (lo == hi) values[lo] else values[lo] + (idx - lo) * (values[hi] - values[lo])
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class QuartileFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val quart = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (quart !in 0..4) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val k = quart / 4.0
        return PercentileFunction().evaluate(
            context,
            currentSheet,
            listOf(args[0], FormulaNode.NumberLiteral(k))
        )
    }
}

internal class StdevFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return when (val collected = collectNumericArgs(context, currentSheet, args)) {
            is CellValue.ErrorValue -> EvalValue.Scalar(collected)
            is List<*> -> {
                val nums = collected.filterIsInstance<Double>()
                if (nums.size < 2) EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
                else EvalValue.Scalar(CellValue.NumberValue(sqrt(sampleVariance(nums))))
            }

            else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class StdevpFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return when (val collected = collectNumericArgs(context, currentSheet, args)) {
            is CellValue.ErrorValue -> EvalValue.Scalar(collected)
            is List<*> -> {
                val nums = collected.filterIsInstance<Double>()
                if (nums.isEmpty()) EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
                else EvalValue.Scalar(CellValue.NumberValue(sqrt(populationVariance(nums))))
            }

            else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class VarFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return when (val collected = collectNumericArgs(context, currentSheet, args)) {
            is CellValue.ErrorValue -> EvalValue.Scalar(collected)
            is List<*> -> {
                val nums = collected.filterIsInstance<Double>()
                if (nums.size < 2) EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
                else EvalValue.Scalar(CellValue.NumberValue(sampleVariance(nums)))
            }

            else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class VarpFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        return when (val collected = collectNumericArgs(context, currentSheet, args)) {
            is CellValue.ErrorValue -> EvalValue.Scalar(collected)
            is List<*> -> {
                val nums = collected.filterIsInstance<Double>()
                if (nums.isEmpty()) EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.DIV_0))
                else EvalValue.Scalar(CellValue.NumberValue(populationVariance(nums)))
            }

            else -> EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }
    }
}

internal class MedianFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = mutableListOf<Double>()
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { values += it } },
            onRange = { value -> rangeNumeric(value)?.let { values += it } }
        )
        if (error != null) return EvalValue.Scalar(error)
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val sorted = values.sorted()
        val mid = sorted.size / 2
        val result = if (sorted.size % 2 == 0) {
            (sorted[mid - 1] + sorted[mid]) / 2.0
        } else {
            sorted[mid]
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class SmallFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val values = collectArrayNumbers(context, currentSheet, args[0], includeDirectConversions = false)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val k = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (k < 1 || k > values.size) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(values.sorted()[k - 1]))
    }
}

internal class LargeFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val values = collectArrayNumbers(context, currentSheet, args[0], includeDirectConversions = false)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (values.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val k = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (k < 1 || k > values.size) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(values.sortedDescending()[k - 1]))
    }
}

internal class ProductFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var product = 1.0
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value -> directNumeric(value)?.let { product *= it } },
            onRange = { value -> rangeNumeric(value)?.let { product *= it } }
        )
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(product))
    }
}

internal class SumSqFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        var sum = 0.0
        val error = forEachArgumentValueWithSource(context, currentSheet, args,
            onDirect = { value ->
                directNumeric(value)?.let { n -> sum += n * n }
            },
            onRange = { value ->
                rangeNumeric(value)?.let { n -> sum += n * n }
            }
        )
        return if (error != null) EvalValue.Scalar(error) else EvalValue.Scalar(CellValue.NumberValue(sum))
    }
}

internal class CeilingFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val significance = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            1.0
        }
        if (significance == 0.0) return EvalValue.Scalar(CellValue.NumberValue(0.0))
        if (number * significance < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val absSig = abs(significance)
        val result = ceil(number / absSig) * absSig
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class FloorFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 1..2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val significance = if (args.size == 2) {
            context.evalScalar(currentSheet, args[1]).asNumberOrNull()
                ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        } else {
            1.0
        }
        if (significance == 0.0) return EvalValue.Scalar(CellValue.NumberValue(0.0))
        if (number * significance < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val absSig = abs(significance)
        val result = floor(number / absSig) * absSig
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class MRoundFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val multiple = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (multiple == 0.0) return EvalValue.Scalar(CellValue.NumberValue(0.0))
        if (number * multiple < 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val m = abs(multiple)
        val result = roundHalfAwayFromZero(number / m, 0) * m
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class SignFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val result = when {
            number > 0.0 -> 1.0
            number < 0.0 -> -1.0
            else -> 0.0
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class PiFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.isNotEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return EvalValue.Scalar(CellValue.NumberValue(PI))
    }
}

internal class PmtFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..5) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val rate = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nper = context.evalScalar(currentSheet, args[1]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pv = context.evalScalar(currentSheet, args[2]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val fv = if (args.size >= 4) context.evalScalar(currentSheet, args[3]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val type = if (args.size >= 5) context.evalScalar(currentSheet, args[4]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        if (nper <= 0.0 || (type != 0.0 && type != 1.0)) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val result = pmt(rate, nper, pv, fv, type)
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class PvFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..5) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val rate = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nper = context.evalScalar(currentSheet, args[1]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pmt = context.evalScalar(currentSheet, args[2]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val fv = if (args.size >= 4) context.evalScalar(currentSheet, args[3]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val type = if (args.size >= 5) context.evalScalar(currentSheet, args[4]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        if (nper <= 0.0 || (type != 0.0 && type != 1.0)) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val result = pv(rate, nper, pmt, fv, type)
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class FvFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..5) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val rate = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nper = context.evalScalar(currentSheet, args[1]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pmt = context.evalScalar(currentSheet, args[2]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pv = if (args.size >= 4) context.evalScalar(currentSheet, args[3]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val type = if (args.size >= 5) context.evalScalar(currentSheet, args[4]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        if (nper <= 0.0 || (type != 0.0 && type != 1.0)) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val result = fv(rate, nper, pmt, pv, type)
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class NperFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..5) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val rate = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pmt = context.evalScalar(currentSheet, args[1]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pv = context.evalScalar(currentSheet, args[2]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val fv = if (args.size >= 4) context.evalScalar(currentSheet, args[3]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val type = if (args.size >= 5) context.evalScalar(currentSheet, args[4]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        if (type != 0.0 && type != 1.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val nper = nper(rate, pmt, pv, fv, type) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(nper))
    }
}

internal class RateFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size !in 3..6) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nper = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pmt = context.evalScalar(currentSheet, args[1]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val pv = context.evalScalar(currentSheet, args[2]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val fv = if (args.size >= 4) context.evalScalar(currentSheet, args[3]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val type = if (args.size >= 5) context.evalScalar(currentSheet, args[4]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.0
        val guess = if (args.size >= 6) context.evalScalar(currentSheet, args[5]).asNumberOrNull() ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE)) else 0.1
        if (nper <= 0.0 || (type != 0.0 && type != 1.0)) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val rate = rate(nper, pmt, pv, fv, type, guess) ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(rate))
    }
}

internal class Log10Function : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)
        if (values.size != 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val number = values.first().asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        return if (number <= 0.0) {
            EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        } else {
            EvalValue.Scalar(CellValue.NumberValue(log10(number)))
        }
    }
}

internal class EffectFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nominalRate = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val npery = context.evalScalar(currentSheet, args[1]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (nominalRate <= 0.0 || npery < 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val result = (1.0 + nominalRate / npery).pow(npery) - 1.0
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class ExponDistFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        if (args.size != 3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val x = context.evalScalar(currentSheet, args[0]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val lambda = context.evalScalar(currentSheet, args[1]).asNumberOrNull()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val cumulative = context.asBoolean(context.evalScalar(currentSheet, args[2]))
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (x < 0.0 || lambda <= 0.0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        val result = if (cumulative) {
            1.0 - exp(-lambda * x)
        } else {
            lambda * exp(-lambda * x)
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class NpvFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)

        if (values.size < 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        val nums = values.mapNotNull { it.asNumberOrNull() }
        if (nums.size < 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))

        val rate = nums[0]
        var result = 0.0
        for (i in 1 until nums.size) {
            result += nums[i] / (1.0 + rate).pow(i)
        }
        return EvalValue.Scalar(CellValue.NumberValue(result))
    }
}

internal class IrrFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)

        val cashflows = values.mapNotNull { it.asNumberOrNull() }
        if (cashflows.size < 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        if (cashflows.none { it > 0.0 } || cashflows.none { it < 0.0 }) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }

        val irr = internalRateOfReturn(cashflows)
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        return EvalValue.Scalar(CellValue.NumberValue(irr))
    }

    private fun internalRateOfReturn(cashflows: List<Double>): Double? {
        fun npv(rate: Double): Double {
            var sum = 0.0
            for (i in cashflows.indices) {
                sum += cashflows[i] / (1.0 + rate).pow(i)
            }
            return sum
        }

        fun derivative(rate: Double): Double {
            var sum = 0.0
            for (i in 1 until cashflows.size) {
                sum += -i * cashflows[i] / (1.0 + rate).pow(i + 1)
            }
            return sum
        }

        var rate = 0.1
        repeat(100) {
            val f = npv(rate)
            val d = derivative(rate)
            if (abs(d) < 1e-12) return@repeat
            val next = rate - (f / d)
            if (!next.isFinite() || next <= -0.999999999) return@repeat
            if (abs(next - rate) < 1e-10) return next
            rate = next
        }

        var low = -0.9999
        var high = 1.0
        var fLow = npv(low)
        var fHigh = npv(high)

        var expand = 0
        while (fLow * fHigh > 0 && expand < 20) {
            high *= 2.0
            fHigh = npv(high)
            expand++
        }

        if (fLow * fHigh > 0) return null

        repeat(200) {
            val mid = (low + high) / 2.0
            val fMid = npv(mid)
            if (abs(fMid) < 1e-10) return mid
            if (fLow * fMid < 0) {
                high = mid
                fHigh = fMid
            } else {
                low = mid
                fLow = fMid
            }
            if (abs(high - low) < 1e-10) return (high + low) / 2.0
        }

        return null
    }
}

internal class PpmtFunction : FormulaFunctionHandler {
    override fun evaluate(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): EvalValue {
        val values = context.collectArgValues(currentSheet, args)
        val firstError = values.firstOrNull { it is CellValue.ErrorValue }
        if (firstError != null) return EvalValue.Scalar(firstError)

        if (values.size !in 4..6) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }

        val numbers = values.mapNotNull { it.asNumberOrNull() }
        if (numbers.size < 4) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
        }

        val rate = numbers[0]
        val per = numbers[1]
        val nper = numbers[2]
        val pv = numbers[3]
        val fv = numbers.getOrElse(4) { 0.0 }
        val type = numbers.getOrElse(5) { 0.0 }

        if (per < 1.0 || per > nper || nper <= 0.0 || (type != 0.0 && type != 1.0)) {
            return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NUM))
        }

        val payment = pmt(rate, nper, pv, fv, type)
        val ipmt = ipmt(rate, per, nper, pv, fv, type, payment)
        val ppmt = payment - ipmt

        return EvalValue.Scalar(CellValue.NumberValue(ppmt))
    }

    private fun pmt(rate: Double, nper: Double, pv: Double, fv: Double, type: Double): Double {
        if (rate == 0.0) {
            return -(pv + fv) / nper
        }

        val factor = (1.0 + rate).pow(nper)
        return -(rate * (fv + pv * factor)) / ((1.0 + rate * type) * (factor - 1.0))
    }

    private fun ipmt(
        rate: Double,
        per: Double,
        nper: Double,
        pv: Double,
        fv: Double,
        type: Double,
        payment: Double
    ): Double {
        if (rate == 0.0) return 0.0
        if (type == 1.0 && per == 1.0) return 0.0

        val temp = (1.0 + rate).pow(per - 1.0)
        var interest = -(pv * temp * rate + payment * (temp - 1.0))
        if (type == 1.0) {
            interest /= (1.0 + rate)
        }
        return interest
    }
}

private fun directNumeric(value: CellValue): Double? = when (value) {
    is CellValue.NumberValue -> value.value
    is CellValue.BooleanValue -> if (value.value) 1.0 else 0.0
    is CellValue.TextValue -> value.value.toDoubleOrNull()
    CellValue.Blank -> null
    is CellValue.ErrorValue -> null
}

private fun rangeNumeric(value: CellValue): Double? = when (value) {
    is CellValue.NumberValue -> value.value
    else -> null
}

private inline fun forEachArgumentValueWithSource(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>,
    onDirect: (CellValue) -> Unit,
    onRange: (CellValue) -> Unit
): CellValue.ErrorValue? {
    for (arg in args) {
        when (val value = context.evalNode(currentSheet, arg)) {
            is EvalValue.Scalar -> {
                if (value.value is CellValue.ErrorValue) return value.value
                onDirect(value.value)
            }

            is EvalValue.Range -> {
                for (cellValue in context.iterateRange(value.sheet, value.start, value.end)) {
                    if (cellValue is CellValue.ErrorValue) return cellValue
                    onRange(cellValue)
                }
            }

            is EvalValue.Range3D -> {
                for (range in value.ranges) {
                    for (cellValue in context.iterateRange(range.sheet, range.start, range.end)) {
                        if (cellValue is CellValue.ErrorValue) return cellValue
                        onRange(cellValue)
                    }
                }
            }
        }
    }
    return null
}

private fun toTextValue(value: CellValue): String = when (value) {
    is CellValue.TextValue -> value.value
    is CellValue.NumberValue -> value.value.toString()
    is CellValue.BooleanValue -> if (value.value) "TRUE" else "FALSE"
    CellValue.Blank -> ""
    is CellValue.ErrorValue -> value.message
}

private fun compareForMatch(left: CellValue, right: CellValue): Int {
    val ln = left.asNumberOrNull()
    val rn = right.asNumberOrNull()
    if (ln != null && rn != null) return ln.compareTo(rn)
    return toTextValue(left).compareTo(toTextValue(right), ignoreCase = false)
}

private fun flattenSingleDimension(
    context: FunctionEvaluationContext,
    currentSheet: String,
    node: FormulaNode
): List<CellValue>? {
    return when (val v = context.evalNode(currentSheet, node)) {
        is EvalValue.Scalar -> listOf(v.value)
        is EvalValue.Range -> context.iterateRange(v.sheet, v.start, v.end).toList()
        is EvalValue.Range3D -> {
            val out = mutableListOf<CellValue>()
            for (r in v.ranges) {
                out += context.iterateRange(r.sheet, r.start, r.end).toList()
            }
            out
        }
    }
}

private fun criteriaPredicate(criteria: CellValue): (CellValue) -> Boolean {
    val raw = toTextValue(criteria)
    val operators = listOf(">=", "<=", "<>", ">", "<", "=")
    val op = operators.firstOrNull { raw.startsWith(it) }
    val tail = if (op != null) raw.removePrefix(op) else raw

    val textPattern = tail
    val numeric = tail.toDoubleOrNull()

    return { value ->
        when {
            op == null -> {
                if (numeric != null) {
                    value.asNumberOrNull()?.let { it == numeric } ?: false
                } else {
                    wildcardMatch(textPattern, toTextValue(value))
                }
            }
            numeric != null -> {
                val n = value.asNumberOrNull()
                if (n == null) {
                    false
                } else {
                    when (op) {
                        ">" -> n > numeric
                        "<" -> n < numeric
                        ">=" -> n >= numeric
                        "<=" -> n <= numeric
                        "=" -> n == numeric
                        "<>" -> n != numeric
                        else -> false
                    }
                }
            }
            else -> {
                val t = toTextValue(value)
                val eq = wildcardMatch(textPattern, t)
                when (op) {
                    "=" -> eq
                    "<>" -> !eq
                    else -> false
                }
            }
        }
    }
}

private fun wildcardMatch(pattern: String, text: String): Boolean {
    val regex = buildString {
        append("^")
        for (ch in pattern) {
            when (ch) {
                '*' -> append(".*")
                '?' -> append(".")
                else -> append(Regex.escape(ch.toString()))
            }
        }
        append("$")
    }.toRegex()
    return regex.matches(text)
}

private data class CriteriaPair(
    val range: EvalValue.Range,
    val predicate: (CellValue) -> Boolean
)

private data class ReferenceBounds(
    val sheet: String,
    val start: CellAddress,
    val end: CellAddress
)

private fun parseCriteriaPairs(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>
): List<CriteriaPair>? {
    if (args.size % 2 != 0) return null
    val out = mutableListOf<CriteriaPair>()
    var i = 0
    while (i < args.size) {
        val range = resolveRangeReference(context, currentSheet, args[i]) ?: return null
        val crit = context.evalScalar(currentSheet, args[i + 1])
        if (crit is CellValue.ErrorValue) return null
        out += CriteriaPair(range, criteriaPredicate(crit))
        i += 2
    }
    return out
}

private fun sameRangeShape(a: EvalValue.Range, b: EvalValue.Range): Boolean {
    val ah = a.end.row - a.start.row
    val aw = a.end.column - a.start.column
    val bh = b.end.row - b.start.row
    val bw = b.end.column - b.start.column
    return ah == bh && aw == bw
}

private fun collectHolidayDates(context: FunctionEvaluationContext, currentSheet: String, node: FormulaNode): Set<LocalDate> {
    val out = mutableSetOf<LocalDate>()
    when (val v = context.evalNode(currentSheet, node)) {
        is EvalValue.Scalar -> {
            v.value.asNumberOrNull()?.let { out += fromExcelSerial(it) }
        }

        is EvalValue.Range -> {
            for (c in context.iterateRange(v.sheet, v.start, v.end)) {
                c.asNumberOrNull()?.let { out += fromExcelSerial(it) }
            }
        }

        is EvalValue.Range3D -> {
            for (r in v.ranges) {
                for (c in context.iterateRange(r.sheet, r.start, r.end)) {
                    c.asNumberOrNull()?.let { out += fromExcelSerial(it) }
                }
            }
        }
    }
    return out
}

private fun evalPairwiseArrayOp(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>,
    op: (Double, Double) -> Double
): EvalValue {
    if (args.size != 2) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    val x = flattenSingleDimension(context, currentSheet, args[0])
        ?.mapNotNull { it.asNumberOrNull() }
        ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    val y = flattenSingleDimension(context, currentSheet, args[1])
        ?.mapNotNull { it.asNumberOrNull() }
        ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    if (x.size != y.size || x.isEmpty()) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.NA))
    var sum = 0.0
    for (i in x.indices) sum += op(x[i], y[i])
    return EvalValue.Scalar(CellValue.NumberValue(sum))
}

private fun collectNumericArgs(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>
): Any {
    val values = mutableListOf<Double>()
    val err = context.forEachArgValue(currentSheet, args) { v ->
        v.asNumberOrNull()?.let { values += it }
    }
    return err ?: values
}

private fun collectArrayNumbers(
    context: FunctionEvaluationContext,
    currentSheet: String,
    node: FormulaNode,
    includeDirectConversions: Boolean
): List<Double>? {
    return when (val value = context.evalNode(currentSheet, node)) {
        is EvalValue.Scalar -> {
            val number = if (includeDirectConversions) {
                directNumeric(value.value)
            } else {
                rangeNumeric(value.value)
            }
            if (number == null) emptyList() else listOf(number)
        }

        is EvalValue.Range -> {
            context.iterateRange(value.sheet, value.start, value.end)
                .mapNotNull { rangeNumeric(it) }
                .toList()
        }

        is EvalValue.Range3D -> {
            val out = mutableListOf<Double>()
            for (range in value.ranges) {
                for (cell in context.iterateRange(range.sheet, range.start, range.end)) {
                    rangeNumeric(cell)?.let { out += it }
                }
            }
            out
        }
    }
}

private fun sampleVariance(nums: List<Double>): Double {
    val mean = nums.average()
    val sumSq = nums.sumOf { d -> (d - mean) * (d - mean) }
    return sumSq / (nums.size - 1)
}

private fun populationVariance(nums: List<Double>): Double {
    val mean = nums.average()
    val sumSq = nums.sumOf { d -> (d - mean) * (d - mean) }
    return sumSq / nums.size
}

private fun resolveRangeReference(
    context: FunctionEvaluationContext,
    currentSheet: String,
    node: FormulaNode
): EvalValue.Range? {
    return when (node) {
        is FormulaNode.CellRef -> {
            val sheet = node.sheet ?: currentSheet
            EvalValue.Range(sheet, node.ref, node.ref)
        }

        is FormulaNode.RangeRef -> {
            val sheet = node.sheet ?: currentSheet
            val start = CellAddress(minOf(node.start.row, node.end.row), minOf(node.start.column, node.end.column))
            val end = CellAddress(maxOf(node.start.row, node.end.row), maxOf(node.start.column, node.end.column))
            EvalValue.Range(sheet, start, end)
        }

        else -> when (val evaluated = context.evalNode(currentSheet, node)) {
            is EvalValue.Range -> evaluated
            is EvalValue.Range3D -> evaluated.ranges.firstOrNull()
            is EvalValue.Scalar -> null
        }
    }
}

private fun resolveReferenceBounds(
    context: FunctionEvaluationContext,
    currentSheet: String,
    node: FormulaNode
): ReferenceBounds? {
    return when (node) {
        is FormulaNode.CellRef -> {
            val sheet = node.sheet ?: currentSheet
            ReferenceBounds(sheet, node.ref, node.ref)
        }

        is FormulaNode.RangeRef -> {
            val sheet = node.sheet ?: currentSheet
            val start = CellAddress(minOf(node.start.row, node.end.row), minOf(node.start.column, node.end.column))
            val end = CellAddress(maxOf(node.start.row, node.end.row), maxOf(node.start.column, node.end.column))
            ReferenceBounds(sheet, start, end)
        }

        else -> when (val evaluated = context.evalNode(currentSheet, node)) {
            is EvalValue.Range -> ReferenceBounds(evaluated.sheet, evaluated.start, evaluated.end)
            is EvalValue.Range3D -> {
                val first = evaluated.ranges.firstOrNull() ?: return null
                ReferenceBounds(first.sheet, first.start, first.end)
            }

            is EvalValue.Scalar -> null
        }
    }
}

private fun roundHalfAwayFromZero(value: Double, digits: Int): Double {
    if (!value.isFinite()) return value
    val factor = 10.0.pow(digits)
    val scaled = value * factor
    val rounded = if (scaled >= 0) floor(scaled + 0.5) else ceil(scaled - 0.5)
    return rounded / factor
}

private fun roundAwayFromZero(value: Double, digits: Int): Double {
    if (!value.isFinite()) return value
    val factor = 10.0.pow(digits)
    val scaled = value * factor
    val rounded = if (scaled >= 0) ceil(scaled) else floor(scaled)
    return rounded / factor
}

private fun roundTowardZero(value: Double, digits: Int): Double {
    if (!value.isFinite()) return value
    val factor = 10.0.pow(digits)
    val scaled = value * factor
    val rounded = if (scaled >= 0) floor(scaled) else ceil(scaled)
    return rounded / factor
}

private fun findOrSearch(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>,
    caseSensitive: Boolean
): EvalValue {
    if (args.size !in 2..3) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    val needleV = context.evalScalar(currentSheet, args[0])
    val haystackV = context.evalScalar(currentSheet, args[1])
    if (needleV is CellValue.ErrorValue) return EvalValue.Scalar(needleV)
    if (haystackV is CellValue.ErrorValue) return EvalValue.Scalar(haystackV)
    val start = if (args.size == 3) {
        context.evalScalar(currentSheet, args[2]).asNumberOrNull()?.toInt()
            ?: return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    } else 1
    if (start < 1) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    val needle = toTextValue(needleV)
    val haystack = toTextValue(haystackV)
    val source = if (caseSensitive) haystack else haystack.lowercase()
    val target = if (caseSensitive) needle else needle.lowercase()
    val idx = source.indexOf(target, startIndex = minOf(start - 1, source.length))
    if (idx < 0) return EvalValue.Scalar(CellValue.ErrorValue(ErrorCodes.VALUE))
    return EvalValue.Scalar(CellValue.NumberValue((idx + 1).toDouble()))
}

private fun fullMonthsBetween(start: LocalDate, end: LocalDate): Int {
    var months = (end.year - start.year) * 12 + (end.monthValue - start.monthValue)
    if (end.dayOfMonth < start.dayOfMonth) months -= 1
    return months
}

private fun fullYearsBetween(start: LocalDate, end: LocalDate): Int {
    var years = end.year - start.year
    if (end.monthValue < start.monthValue || (end.monthValue == start.monthValue && end.dayOfMonth < start.dayOfMonth)) {
        years -= 1
    }
    return years
}

private fun datedifMd(start: LocalDate, end: LocalDate): Int {
    val shifted = start.plusMonths(fullMonthsBetween(start, end).toLong())
    return ChronoUnit.DAYS.between(shifted, end).toInt()
}

private fun days360(startDate: LocalDate, endDate: LocalDate, usMethod: Boolean): Int {
    var sDay = startDate.dayOfMonth
    var eDay = endDate.dayOfMonth
    val sMonth = startDate.monthValue
    val eMonth = endDate.monthValue
    val sYear = startDate.year
    val eYear = endDate.year

    if (usMethod) {
        if (sDay == 31) sDay = 30
        if (eDay == 31 && sDay >= 30) eDay = 30
    } else {
        if (sDay == 31) sDay = 30
        if (eDay == 31) eDay = 30
    }

    return (eYear - sYear) * 360 + (eMonth - sMonth) * 30 + (eDay - sDay)
}

private fun extractTimeFraction(
    context: FunctionEvaluationContext,
    currentSheet: String,
    args: List<FormulaNode>
): Double? {
    if (args.size != 1) return null
    val value = context.evalScalar(currentSheet, args[0]).asNumberOrNull() ?: return null
    val fraction = value - floor(value)
    return if (fraction >= 0.0) fraction else (fraction + 1.0)
}

private fun parseTimeText(input: String): LocalTime? {
    val candidates = listOf(
        DateTimeFormatter.ofPattern("H:mm"),
        DateTimeFormatter.ofPattern("H:mm:ss"),
        DateTimeFormatter.ofPattern("H:mm:ss.SSS"),
        DateTimeFormatter.ofPattern("h:mm a"),
        DateTimeFormatter.ofPattern("h:mm:ss a")
    )
    for (fmt in candidates) {
        try {
            return LocalTime.parse(input.uppercase(Locale.US), fmt)
        } catch (_: DateTimeParseException) {
        }
    }
    return null
}

private fun applyTextFormat(number: Double, pattern: String): String {
    val normalized = pattern.trim().lowercase(Locale.US)
    return when (normalized) {
        "0", "0.0", "0.00", "#,##0", "#,##0.0", "#,##0.00" -> {
            val decimals = normalized.substringAfter('.', "").length
            val grouping = normalized.contains(",")
            formatNumber(number, decimals, grouping)
        }

        "yyyy-mm-dd" -> {
            val date = fromExcelSerial(number)
            date.format(DateTimeFormatter.ISO_LOCAL_DATE)
        }

        "hh:mm:ss" -> {
            val seconds = (((number - floor(number) + 1.0) % 1.0) * 86400.0).toLong()
            val h = (seconds / 3600).toInt()
            val m = ((seconds % 3600) / 60).toInt()
            val s = (seconds % 60).toInt()
            "%02d:%02d:%02d".format(Locale.US, h, m, s)
        }

        else -> number.toString()
    }
}

private fun formatNumber(number: Double, decimals: Int, useGrouping: Boolean): String {
    val safeDecimals = max(0, decimals)
    val pattern = buildString {
        append(if (useGrouping) "#,##0" else "0")
        if (safeDecimals > 0) {
            append(".")
            repeat(safeDecimals) { append("0") }
        }
    }
    val symbols = DecimalFormatSymbols(Locale.US)
    val df = DecimalFormat(pattern, symbols)
    df.isGroupingUsed = useGrouping
    return df.format(number)
}

private fun gcdLong(a: Long, b: Long): Long {
    var x = kotlin.math.abs(a)
    var y = kotlin.math.abs(b)
    while (y != 0L) {
        val t = x % y
        x = y
        y = t
    }
    return x
}

private fun extractDateArg(context: FunctionEvaluationContext, currentSheet: String, args: List<FormulaNode>): LocalDate? {
    if (args.size != 1) return null
    val serial = context.evalScalar(currentSheet, args[0]).asNumberOrNull()?.toLong() ?: return null
    return fromExcelSerial(serial.toDouble())
}

private val excelEpoch: LocalDate = LocalDate.of(1899, 12, 30)

private fun toExcelSerial(date: LocalDate): Double {
    return ChronoUnit.DAYS.between(excelEpoch, date).toDouble()
}

private fun fromExcelSerial(serial: Double): LocalDate {
    return excelEpoch.plusDays(floor(serial).toLong())
}

private fun pmt(rate: Double, nper: Double, pv: Double, fv: Double, type: Double): Double {
    if (nper == 0.0) return Double.NaN
    if (rate == 0.0) return -(pv + fv) / nper
    val factor = (1.0 + rate).pow(nper)
    return -(rate * (fv + pv * factor)) / ((1.0 + rate * type) * (factor - 1.0))
}

private fun pv(rate: Double, nper: Double, pmt: Double, fv: Double, type: Double): Double {
    return if (rate == 0.0) {
        -(fv + pmt * nper)
    } else {
        val factor = (1.0 + rate).pow(nper)
        -((fv + pmt * (1.0 + rate * type) * (factor - 1.0) / rate) / factor)
    }
}

private fun fv(rate: Double, nper: Double, pmt: Double, pv: Double, type: Double): Double {
    return if (rate == 0.0) {
        -(pv + pmt * nper)
    } else {
        val factor = (1.0 + rate).pow(nper)
        -(pv * factor + pmt * (1.0 + rate * type) * (factor - 1.0) / rate)
    }
}

private fun nper(rate: Double, pmt: Double, pv: Double, fv: Double, type: Double): Double? {
    if (rate == 0.0) {
        if (pmt == 0.0) return null
        return -(pv + fv) / pmt
    }
    val num = pmt * (1.0 + rate * type) - fv * rate
    val den = pv * rate + pmt * (1.0 + rate * type)
    if (num == 0.0 || den == 0.0) return null
    val ratio = num / den
    if (!ratio.isFinite() || ratio <= 0.0) return null
    return ln(ratio) / ln(1.0 + rate)
}

private fun rate(nper: Double, pmt: Double, pv: Double, fv: Double, type: Double, guess: Double): Double? {
    fun f(r: Double): Double {
        return if (r == 0.0) {
            pv + pmt * nper + fv
        } else {
            val factor = (1.0 + r).pow(nper)
            pv * factor + pmt * (1.0 + r * type) * (factor - 1.0) / r + fv
        }
    }
    var r = guess
    repeat(100) {
        val fr = f(r)
        val h = 1e-7
        val d = (f(r + h) - fr) / h
        if (!d.isFinite() || abs(d) < 1e-12) return@repeat
        val next = r - fr / d
        if (!next.isFinite()) return@repeat
        if (abs(next - r) < 1e-12) return next
        r = next
    }
    return null
}
