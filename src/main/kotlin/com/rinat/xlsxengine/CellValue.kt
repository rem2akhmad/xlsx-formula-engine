package com.rinat.xlsxengine

sealed interface CellValue {
    data class NumberValue(val value: Double) : CellValue
    data class TextValue(val value: String) : CellValue
    data class BooleanValue(val value: Boolean) : CellValue
    data object Blank : CellValue
    data class ErrorValue(val message: String) : CellValue
}

fun CellValue.asNumberOrNull(): Double? = when (this) {
    is CellValue.NumberValue -> value
    is CellValue.BooleanValue -> if (value) 1.0 else 0.0
    is CellValue.TextValue -> value.toDoubleOrNull()
    CellValue.Blank -> 0.0
    is CellValue.ErrorValue -> null
}
