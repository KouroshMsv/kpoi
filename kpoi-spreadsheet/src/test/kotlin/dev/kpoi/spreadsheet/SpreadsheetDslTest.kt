package dev.kpoi.spreadsheet

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertNotNull
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream
import java.time.LocalDate

class SpreadsheetDslTest {

    @Test
    fun `round-trips values, styles, formulas, and merges`() {
        val bytes = workbook {
            val header = style {
                font {
                    bold = true
                    color("#FFFFFF")
                }
                fill("#4472C4")
                align(horizontal = HorizontalAlignment.CENTER)
            }
            sheet("Report") {
                columnWidth(0, 20)
                freeze(rows = 1)
                row(header) {
                    cell("Name")
                    cell("Score")
                }
                row {
                    cell("Alice")
                    cell(95.5) { format = "0.00" }
                }
                row {
                    cell("Bob")
                    cell(87.25) { format = "0.00" }
                }
                row {
                    cell("Total")
                    formula("SUM(B2:B3)")
                }
                merge("A6:B6")
                autoFilter("A1:B1")
            }
        }.use { it.toByteArray() }

        XSSFWorkbook(ByteArrayInputStream(bytes)).use { workbook ->
            val sheet = workbook.getSheet("Report")
            assertNotNull(sheet)
            assertEquals("Name", sheet.cellAt("A1")?.stringOrNull())
            assertEquals("Alice", sheet["A2"]?.stringOrNull())
            assertEquals(95.5, sheet.cellAt("B2")?.doubleOrNull())
            assertEquals("SUM(B2:B3)", sheet.getRow(3).getCell(1).cellFormula)
            assertEquals(1, sheet.numMergedRegions)
            assertEquals("0.00", sheet.cellAt("B2")?.cellStyle?.dataFormatString)

            val headerStyle = sheet.cellAt("A1")?.cellStyle as XSSFCellStyle
            assertTrue(headerStyle.font.bold)
            assertEquals(HorizontalAlignment.CENTER, headerStyle.alignment)
            val fill = headerStyle.fillForegroundColorColor as XSSFColor
            assertEquals("FF4472C4", fill.getARGBHex())
        }
    }

    @Test
    fun `deduplicates identical styles across cells`() {
        workbook {
            sheet {
                repeat(200) { rowIndex ->
                    row {
                        repeat(5) { columnIndex ->
                            cell("value $rowIndex:$columnIndex") {
                                style {
                                    font { bold = true }
                                }
                            }
                        }
                    }
                }
            }
        }.use { workbook ->
            // 1000 styled cells share exactly one new style and font
            // beyond the workbook defaults.
            assertEquals(2, workbook.numCellStyles)
            assertEquals(2, workbook.numberOfFonts)
        }
    }

    @Test
    fun `date cells get the default date format`() {
        workbook {
            sheet {
                row { cell(LocalDate.of(2026, 7, 19)) }
            }
        }.use { workbook ->
            val cell = workbook.getSheetAt(0).getRow(0).getCell(0)
            assertEquals(LocalDate.of(2026, 7, 19), cell.localDateOrNull())
            assertEquals("yyyy-mm-dd", cell.cellStyle.dataFormatString)
        }
    }

    @Test
    fun `hyperlink cells carry link and label`() {
        workbook {
            sheet {
                row { hyperlink("https://poi.apache.org", "POI") }
            }
        }.use { workbook ->
            val cell = workbook.getSheetAt(0).getRow(0).getCell(0)
            assertEquals("POI", cell.stringCellValue)
            assertEquals("https://poi.apache.org", cell.hyperlink.address)
        }
    }

    @Test
    fun `xls round-trips with indexed colors`() {
        val bytes = xlsWorkbook {
            val header = style {
                fill(IndexedColors.YELLOW)
                font { bold = true }
            }
            sheet("Legacy") {
                row(header) { cell("A") }
            }
        }.use { it.toByteArray() }

        HSSFWorkbook(ByteArrayInputStream(bytes)).use { workbook ->
            val cell = workbook.getSheetAt(0).getRow(0).getCell(0)
            assertEquals("A", cell.stringCellValue)
            assertEquals(IndexedColors.YELLOW.index, cell.cellStyle.fillForegroundColor)
        }
    }

    @Test
    fun `hex colors on xls fail with a helpful message`() {
        val error = assertThrows(UnsupportedOperationException::class.java) {
            xlsWorkbook {
                val red = style { fill("#FF0000") }
                sheet {
                    row(red) { cell("x") }
                }
            }
        }
        assertTrue(error.message!!.contains("IndexedColors"))
    }

    @Test
    fun `streaming workbook writes many rows`() {
        val bytes = streamingWorkbook {
            sheet("Big") {
                repeat(5000) { index ->
                    row {
                        cell("row $index")
                        cell(index)
                    }
                }
            }
        }.use { it.toByteArray() }

        XSSFWorkbook(ByteArrayInputStream(bytes)).use { workbook ->
            assertEquals(5000, workbook.getSheetAt(0).physicalNumberOfRows)
            assertEquals("row 4999", workbook.getSheetAt(0).getRow(4999).getCell(0).stringCellValue)
        }
    }

    @Test
    fun `row styles merge with cell overrides`() {
        workbook {
            val base = style {
                font { bold = true }
                fill("#DDDDDD")
            }
            sheet {
                row(base) {
                    cell("plain")
                    cell("red") {
                        style {
                            font { color("#FF0000") }
                        }
                    }
                }
            }
        }.use { workbook ->
            val row = workbook.getSheetAt(0).getRow(0)
            val plain = row.getCell(0).cellStyle as XSSFCellStyle
            val red = row.getCell(1).cellStyle as XSSFCellStyle
            assertTrue(plain.font.bold)
            assertTrue(red.font.bold)
            assertEquals("FFFF0000", red.font.xssfColor.getARGBHex())
            // both inherit the same fill from the row style
            assertEquals(
                (plain.fillForegroundColorColor as XSSFColor).getARGBHex(),
                (red.fillForegroundColorColor as XSSFColor).getARGBHex(),
            )
        }
    }
}
