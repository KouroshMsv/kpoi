package io.github.kouroshmsv.kpoi.word

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFTable

/** Top-level DSL scope for building a Word document. */
@PoiDsl
public class DocumentBuilder internal constructor(
    /** The underlying POI document, for anything the DSL does not cover. */
    public val poiDocument: XWPFDocument,
) {
    /** Adds a paragraph; [text] becomes its first run. */
    public fun paragraph(text: String? = null, block: ParagraphBuilder.() -> Unit = {}): XWPFParagraph {
        val poiParagraph = poiDocument.createParagraph()
        val builder = ParagraphBuilder(poiParagraph)
        if (text != null) {
            builder.text(text)
        }
        builder.block()
        builder.applyProperties()
        return poiParagraph
    }

    /**
     * Adds a heading rendered as bold text sized by [level] 1..4
     * (20 / 16 / 14 / 12 pt). Real Word heading styles are on the roadmap.
     */
    public fun heading(text: String, level: Int = 1, block: RunBuilder.() -> Unit = {}): XWPFParagraph {
        val headingSize = when (level.coerceIn(1, 4)) {
            1 -> 20
            2 -> 16
            3 -> 14
            else -> 12
        }
        return paragraph {
            text(text) {
                bold = true
                size = headingSize
                block()
            }
        }
    }

    /** Starts a new page. */
    public fun pageBreak() {
        poiDocument.createParagraph().isPageBreak = true
    }

    /** Adds a table; [width] accepts `"100%"`, `"auto"`, or twips like `"5000"`. */
    public fun table(width: String? = null, block: TableBuilder.() -> Unit): XWPFTable {
        val poiTable = poiDocument.createTable()
        width?.let { poiTable.setWidth(it) }
        TableBuilder(poiTable).block()
        return poiTable
    }
}
