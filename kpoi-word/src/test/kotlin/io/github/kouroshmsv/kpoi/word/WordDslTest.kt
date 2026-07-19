package io.github.kouroshmsv.kpoi.word

import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream

class WordDslTest {

    @Test
    fun `round-trips paragraphs, runs, headings, and tables`() {
        val bytes = document {
            heading("Report", level = 1)
            paragraph {
                text("Hello, ")
                text("world") {
                    bold = true
                    color = "#FF0000"
                }
            }
            paragraph("Centered") {
                alignment = ParagraphAlignment.CENTER
            }
            table(width = "100%") {
                row {
                    cell("H1")
                    cell("H2")
                }
                row {
                    cell("a")
                    cell("b")
                }
            }
            pageBreak()
            paragraph("After break")
        }.use { it.toByteArray() }

        XWPFDocument(ByteArrayInputStream(bytes)).use { document ->
            val paragraphs = document.paragraphs
            assertEquals("Report", paragraphs[0].text)
            assertTrue(paragraphs[0].runs[0].isBold)
            assertEquals(20.0, paragraphs[0].runs[0].fontSizeAsDouble)

            assertEquals("Hello, world", paragraphs[1].text)
            val styled = paragraphs[1].runs[1]
            assertTrue(styled.isBold)
            assertEquals("FF0000", styled.color)

            assertEquals(ParagraphAlignment.CENTER, paragraphs[2].alignment)

            val table = document.tables[0]
            assertEquals("H2", table.getRow(0).getCell(1).text)
            assertEquals("a", table.getRow(1).getCell(0).text)

            assertTrue(paragraphs[3].isPageBreak)
            assertEquals("After break", paragraphs[4].text)
        }
    }

    @Test
    fun `multi-line text becomes runs with breaks`() {
        document {
            paragraph("first\nsecond")
        }.use { document ->
            // both lines live in one run, separated by a <w:br/>
            assertEquals("first\nsecond", document.paragraphs[0].runs[0].text().replace("\r", ""))
        }
    }
}
