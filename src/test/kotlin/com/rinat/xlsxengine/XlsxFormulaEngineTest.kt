package com.rinat.xlsxengine

import java.nio.file.Files
import java.nio.file.Path
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertTrue

class XlsxFormulaEngineTest {
    @Test
    fun `evaluates formulas from xlsx package`() {
        val temp = Files.createTempFile("xlsx-engine-", ".xlsx")
        writeFixtureWorkbook(temp)

        val engine = XlsxFormulaEngine.fromFile(temp)

        val a3 = engine.evaluateCell("Sheet1", "A3")
        assertEquals(CellValue.NumberValue(30.0), a3)

        val b1 = engine.evaluateCell("Sheet1", "B1")
        assertEquals(CellValue.NumberValue(60.0), b1)

        val b2 = engine.evaluateCell("Sheet1", "B2")
        assertEquals(CellValue.NumberValue(20.0), b2)

        val c1 = engine.evaluateCell("Sheet1", "C1")
        assertEquals(CellValue.TextValue("hello"), c1)

        val c2 = engine.evaluateCell("Sheet1", "C2")
        assertEquals(CellValue.NumberValue(2.0), c2)

        val sheet2A1 = engine.evaluateCell("Sheet2", "A1")
        assertEquals(CellValue.NumberValue(65.0), sheet2A1)

        val c3 = engine.evaluateCell("Sheet1", "C3")
        assertTrue(c3 is CellValue.ErrorValue)

        val row2 = engine.evaluateRow("Sheet1", 2)
        assertEquals(CellValue.NumberValue(20.0), row2["A2"])
        assertEquals(CellValue.NumberValue(20.0), row2["B2"])

        val columnA = engine.evaluateColumn("Sheet1", "A")
        assertEquals(CellValue.NumberValue(10.0), columnA[1])
        assertEquals(CellValue.NumberValue(20.0), columnA[2])
        assertEquals(CellValue.NumberValue(30.0), columnA[3])
    }

    private fun writeFixtureWorkbook(path: Path) {
        ZipOutputStream(Files.newOutputStream(path)).use { zip ->
            putEntry(
                zip,
                "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                </Types>
                """.trimIndent()
            )

            putEntry(
                zip,
                "xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
                  </sheets>
                </workbook>
                """.trimIndent()
            )

            putEntry(
                zip,
                "xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
                </Relationships>
                """.trimIndent()
            )

            putEntry(
                zip,
                "xl/sharedStrings.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
                  <si><t>hello</t></si>
                </sst>
                """.trimIndent()
            )

            putEntry(
                zip,
                "xl/worksheets/sheet1.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1"><v>10</v></c>
                      <c r="B1"><f>A3*2</f><v>0</v></c>
                      <c r="C1" t="s"><v>0</v></c>
                    </row>
                    <row r="2">
                      <c r="A2"><v>20</v></c>
                      <c r="B2"><f>AVERAGE(A1:A2,30)</f><v>0</v></c>
                      <c r="C2"><f>COUNT(A1:A2,C1)</f><v>0</v></c>
                    </row>
                    <row r="3">
                      <c r="A3"><f>SUM(A1:A2)</f><v>0</v></c>
                      <c r="C3"><f>A1/A4</f><v>0</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
            )

            putEntry(
                zip,
                "xl/worksheets/sheet2.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1"><f>Sheet1!B1+5</f><v>0</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
            )
        }
    }

    private fun putEntry(zip: ZipOutputStream, name: String, content: String) {
        zip.putNextEntry(ZipEntry(name))
        zip.write(content.toByteArray(Charsets.UTF_8))
        zip.closeEntry()
    }
}
