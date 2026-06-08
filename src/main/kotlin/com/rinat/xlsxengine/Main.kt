package com.rinat.xlsxengine

import java.nio.file.Paths

fun main() {
    val filePath = Paths.get("src/test/resources/Savings modeling.xlsx")
    val engine = XlsxFormulaEngine.fromFile(filePath)
    val workbook = XlsxReader.read(filePath)
    val sheetName = workbook.sheetOrder.firstOrNull { it.equals("sheet3", ignoreCase = true) }
        ?: error("Sheet 'sheet3' not found. Available: ${workbook.sheetOrder}")

    val before = engine.evaluateCell(sheetName, "F4")
    println("Before: $sheetName!F4 = $before")

    engine.setCellNumber("Sheet3", "C4", 5.0)
    val after = engine.evaluateCell(sheetName, "F4")
    println("After:  $sheetName!F4 = $after")
}
