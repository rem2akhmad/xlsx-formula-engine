package com.rinat.xlsxengine

internal sealed interface EvalValue {
    data class Scalar(val value: CellValue) : EvalValue

    data class Range(
        val sheet: String,
        val start: CellAddress,
        val end: CellAddress
    ) : EvalValue

    data class Range3D(
        val ranges: List<Range>
    ) : EvalValue
}
