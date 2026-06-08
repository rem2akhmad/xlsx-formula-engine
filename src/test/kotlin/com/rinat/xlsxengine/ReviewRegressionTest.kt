package com.rinat.xlsxengine

import java.nio.file.Files
import java.nio.file.Path
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertFalse
import kotlin.test.assertTrue

class ReviewRegressionTest {
    @Test
    fun `evaluateRow should use current workbook after mutations`() {
        val temp = Files.createTempFile("review-row-", ".xlsx")
        writeWorkbook(
            temp,
            sheetXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1"><v>1</v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
        )

        val engine = XlsxFormulaEngine.fromFile(temp)
        val rowBefore = engine.evaluateRow("Sheet1", 1)
        assertTrue("A1" in rowBefore)
        assertFalse("B1" in rowBefore)

        engine.setCellNumber("Sheet1", "B1", 5.0)
        val rowAfterAdd = engine.evaluateRow("Sheet1", 1)
        assertTrue("B1" in rowAfterAdd)

        engine.clearCell("Sheet1", "A1")
        val rowAfterClear = engine.evaluateRow("Sheet1", 1)
        assertFalse("A1" in rowAfterClear)
        assertTrue("B1" in rowAfterClear)
    }

    @Test
    fun `COUNT should not include blanks in ranges`() {
        val temp = Files.createTempFile("review-count-", ".xlsx")
        writeWorkbook(
            temp,
            sheetXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1"><v>10</v></c>
                      <c r="B1"><f>COUNT(A1:A3)</f><v>0</v></c>
                    </row>
                    <row r="2">
                      <c r="A2"/>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
        )

        val engine = XlsxFormulaEngine.fromFile(temp)
        assertEquals(CellValue.NumberValue(1.0), engine.evaluateCell("Sheet1", "B1"))
    }

    @Test
    fun `inlineStr rich text runs should be concatenated`() {
        val temp = Files.createTempFile("review-inline-", ".xlsx")
        writeWorkbook(
            temp,
            sheetXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="inlineStr">
                        <is>
                          <r><t>Hello </t></r>
                          <r><t>World</t></r>
                        </is>
                      </c>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
        )

        val engine = XlsxFormulaEngine.fromFile(temp)
        assertEquals(CellValue.TextValue("Hello World"), engine.evaluateCell("Sheet1", "A1"))
    }

    @Test
    fun `v text value should preserve meaningful spaces`() {
        val temp = Files.createTempFile("review-v-space-", ".xlsx")
        writeWorkbook(
            temp,
            sheetXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1">
                      <c r="A1" t="str"><v>  padded value  </v></c>
                    </row>
                  </sheetData>
                </worksheet>
                """.trimIndent()
        )

        val engine = XlsxFormulaEngine.fromFile(temp)
        assertEquals(CellValue.TextValue("  padded value  "), engine.evaluateCell("Sheet1", "A1"))
    }

    private fun writeWorkbook(path: Path, sheetXml: String) {
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
                </Relationships>
                """.trimIndent()
            )
            putEntry(zip, "xl/worksheets/sheet1.xml", sheetXml)
        }
    }

    private fun putEntry(zip: ZipOutputStream, name: String, content: String) {
        zip.putNextEntry(ZipEntry(name))
        zip.write(content.toByteArray(Charsets.UTF_8))
        zip.closeEntry()
    }
}
