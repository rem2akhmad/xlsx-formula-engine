package com.rinat.xlsxengine

data class WorkbookCell(
    val address: CellAddress,
    val type: String?,
    val rawValue: String?,
    val formula: String?
)

data class WorksheetData(
    val name: String,
    val cells: Map<CellAddress, WorkbookCell>
)

data class WorkbookData(
    val sheets: Map<String, WorksheetData>,
    val sheetOrder: List<String>,
    val sharedStrings: List<String>,
    val calcChain: List<CalcChainEntry> = emptyList(),
    val definedNames: Map<String, String> = emptyMap(),
    val localDefinedNames: Map<String, Map<String, String>> = emptyMap()
)

data class CalcChainEntry(
    val sheet: String,
    val address: CellAddress
)

data class EvaluatedWorkbook(
    val sheets: Map<String, EvaluatedSheet>
)

data class EvaluatedSheet(
    val name: String,
    val values: Map<CellAddress, CellValue>
)
