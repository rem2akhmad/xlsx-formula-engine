package com.rinat.xlsxengine

import java.nio.file.Files
import java.nio.file.Path
import java.util.zip.ZipEntry
import java.util.zip.ZipOutputStream
import kotlin.test.Test
import kotlin.test.assertEquals
import kotlin.test.assertIs

class SpecificationComplianceRegressionTest {
    @Test
    fun `supports postfix percent and concat operators`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1"><v>20</v></c>
                          <c r="B1"><f>A1*10%</f></c>
                          <c r="C1"><f>"Rate: "&amp;B1</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(2.0), engine.evaluateCell("Sheet1", "B1"))
        assertEquals(CellValue.TextValue("Rate: 2.0"), engine.evaluateCell("Sheet1", "C1"))
    }

    @Test
    fun `supports 3D ranges`() {
        val path = createWorkbook(
            workbookXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
                    <sheet name="Sheet3" sheetId="3" r:id="rId3"/>
                    <sheet name="Sheet4" sheetId="4" r:id="rId4"/>
                  </sheets>
                </workbook>
                """.trimIndent(),
            relsXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
                  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
                </Relationships>
                """.trimIndent(),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to singleCellSheetXml("A1", "1"),
                "xl/worksheets/sheet2.xml" to singleCellSheetXml("A1", "2"),
                "xl/worksheets/sheet3.xml" to singleCellSheetXml("A1", "3"),
                "xl/worksheets/sheet4.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1"><f>SUM(Sheet1:Sheet3!A1)</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(6.0), engine.evaluateCell("Sheet4", "A1"))
    }

    @Test
    fun `MIN and MAX return zero when no numeric values found`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1" t="str"><v>x</v></c>
                          <c r="B1" t="b"><v>1</v></c>
                          <c r="C1"><f>MIN(A1)</f></c>
                          <c r="D1"><f>MAX(A1)</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(0.0), engine.evaluateCell("Sheet1", "C1"))
        assertEquals(CellValue.NumberValue(0.0), engine.evaluateCell("Sheet1", "D1"))
    }

    @Test
    fun `SUM COUNT AVERAGE use different coercion for direct args vs ranges`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1" t="str"><v>5</v></c>
                          <c r="B1" t="b"><v>1</v></c>
                          <c r="C1"><v>3</v></c>
                        </row>
                        <row r="2">
                          <c r="A2"><f>SUM("1",2)</f></c>
                          <c r="B2"><f>SUM(A1:C1)</f></c>
                          <c r="C2"><f>COUNT(TRUE,1)</f></c>
                        </row>
                        <row r="3">
                          <c r="A3"><f>COUNT(A1:C1)</f></c>
                          <c r="B3"><f>AVERAGE(TRUE,3)</f></c>
                          <c r="C3"><f>AVERAGE(B1:C1)</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(3.0), engine.evaluateCell("Sheet1", "A2"))
        assertEquals(CellValue.NumberValue(3.0), engine.evaluateCell("Sheet1", "B2"))
        assertEquals(CellValue.NumberValue(2.0), engine.evaluateCell("Sheet1", "C2"))
        assertEquals(CellValue.NumberValue(1.0), engine.evaluateCell("Sheet1", "A3"))
        assertEquals(CellValue.NumberValue(2.0), engine.evaluateCell("Sheet1", "B3"))
        assertEquals(CellValue.NumberValue(3.0), engine.evaluateCell("Sheet1", "C3"))
    }

    @Test
    fun `returns standard error code for cycles`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1"><f>B1</f></c>
                          <c r="B1"><f>A1</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.ErrorValue(ErrorCodes.REF), engine.evaluateCell("Sheet1", "A1"))
    }

    @Test
    fun `supports simple defined names and R1C1 references`() {
        val path = createWorkbook(
            workbookXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <definedNames>
                    <definedName name="BaseRate">Sheet1!${'$'}A${'$'}1</definedName>
                  </definedNames>
                  <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                  </sheets>
                </workbook>
                """.trimIndent(),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="A1"><v>10</v></c>
                          <c r="B1"><f>BaseRate+R1C1</f></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(20.0), engine.evaluateCell("Sheet1", "B1"))
    }

    @Test
    fun `shared formula shift outside sheet bounds yields REF error`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1">
                          <c r="B1"><f t="shared" si="0">A1</f></c>
                          <c r="A1"><f t="shared" si="0"/></c>
                        </row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            )
        )

        val workbook = XlsxReader.read(path)
        val shiftedFormula = workbook.sheets.getValue("Sheet1").cells.getValue(CellAddress.parse("A1")).formula
        assertEquals("#REF!", shiftedFormula)

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.ErrorValue(ErrorCodes.REF), engine.evaluateCell("Sheet1", "A1"))
    }

    @Test
    fun `resolves worksheet targets using normalized OPC relative paths`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/../worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to singleCellSheetXml("A1", "7")
            )
        )

        val engine = XlsxFormulaEngine.fromFile(path)
        assertEquals(CellValue.NumberValue(7.0), engine.evaluateCell("Sheet1", "A1"))
    }

    @Test
    fun `reads calcChain entries`() {
        val path = createWorkbook(
            workbookXml = singleSheetWorkbookXml("Sheet1", "rId1"),
            relsXml = singleSheetRelsXml("rId1", "worksheets/sheet1.xml"),
            sheetEntries = mapOf(
                "xl/worksheets/sheet1.xml" to
                    """
                    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                      <sheetData>
                        <row r="1"><c r="A1"><v>1</v></c></row>
                      </sheetData>
                    </worksheet>
                    """.trimIndent()
            ),
            calcChainXml =
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <c r="A1" i="1"/>
                </calcChain>
                """.trimIndent()
        )

        val workbook = XlsxReader.read(path)
        assertEquals(1, workbook.calcChain.size)
        val entry = workbook.calcChain.first()
        assertEquals("Sheet1", entry.sheet)
        assertEquals(CellAddress.parse("A1"), entry.address)

        val engine = XlsxFormulaEngine.fromFile(path)
        val all = engine.evaluateAll()
        val cell = all.sheets.getValue("Sheet1").values.getValue(CellAddress.parse("A1"))
        assertIs<CellValue.NumberValue>(cell)
    }

    private fun createWorkbook(
        workbookXml: String,
        relsXml: String,
        sheetEntries: Map<String, String>,
        calcChainXml: String? = null
    ): Path {
        val temp = Files.createTempFile("spec-compliance-", ".xlsx")
        ZipOutputStream(Files.newOutputStream(temp)).use { zip ->
            putEntry(
                zip,
                "[Content_Types].xml",
                buildString {
                    appendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
                    appendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">")
                    appendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>")
                    appendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>")
                    appendLine("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>")
                    for (sheetPath in sheetEntries.keys.sorted()) {
                        appendLine("  <Override PartName=\"/${sheetPath}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>")
                    }
                    if (calcChainXml != null) {
                        appendLine("  <Override PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>")
                    }
                    append("</Types>")
                }
            )
            putEntry(zip, "xl/workbook.xml", workbookXml)
            putEntry(zip, "xl/_rels/workbook.xml.rels", relsXml)
            for ((name, content) in sheetEntries) {
                putEntry(zip, name, content)
            }
            if (calcChainXml != null) {
                putEntry(zip, "xl/calcChain.xml", calcChainXml)
            }
        }
        return temp
    }

    private fun singleSheetWorkbookXml(sheetName: String, relId: String): String =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="$sheetName" sheetId="1" r:id="$relId"/>
          </sheets>
        </workbook>
        """.trimIndent()

    private fun singleSheetRelsXml(relId: String, target: String): String =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="$relId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="$target"/>
        </Relationships>
        """.trimIndent()

    private fun singleCellSheetXml(address: String, value: String): String =
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="$address"><v>$value</v></c>
            </row>
          </sheetData>
        </worksheet>
        """.trimIndent()

    private fun putEntry(zip: ZipOutputStream, name: String, content: String) {
        zip.putNextEntry(ZipEntry(name))
        zip.write(content.toByteArray(Charsets.UTF_8))
        zip.closeEntry()
    }
}
