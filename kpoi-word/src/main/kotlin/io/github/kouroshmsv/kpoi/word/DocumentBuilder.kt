package io.github.kouroshmsv.kpoi.word

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFTable

/**
 * Top-level DSL scope for building a Word document, and the receiver of the
 * [document] and [docx] builder lambdas.
 *
 * From here you add [paragraph]s, [heading]s, and [table]s, and insert
 * [pageBreak]s. The wrapped POI document is exposed as [poiDocument] for
 * anything the DSL does not cover.
 */
@PoiDsl
public class DocumentBuilder internal constructor(
    /** The underlying POI document, for anything the DSL does not cover. */
    public val poiDocument: XWPFDocument,
) {
    /**
     * Adds a paragraph to the document.
     *
     * When [text] is given it becomes the paragraph's first run; further runs
     * and paragraph-level properties (alignment, spacing) are configured in
     * [block] against a [ParagraphBuilder].
     *
     * ```kotlin
     * paragraph("Centered") {
     *     alignment = ParagraphAlignment.CENTER
     *     text(" and bold") { bold = true }
     * }
     * ```
     *
     * @param text optional text for the first run; omit for an empty or
     *   run-by-run paragraph.
     * @param block configures the paragraph and its runs.
     * @return the created [XWPFParagraph].
     */
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
     * Adds a heading rendered as a single bold run whose size is chosen by
     * [level]: `1` -> 20 pt, `2` -> 16 pt, `3` -> 14 pt, `4` -> 12 pt. Levels
     * outside `1..4` are coerced into that range. This is deliberately a visual
     * heading; real Word heading styles are on the roadmap.
     *
     * ```kotlin
     * heading("Quarterly Letter")         // level 1: bold, 20 pt
     * heading("Details", level = 2) {     // bold, 16 pt, plus:
     *     color = "#C00000"
     * }
     * ```
     *
     * @param text heading text.
     * @param level heading level; sizes are 20/16/14/12 pt for 1/2/3/4, with
     *   out-of-range values clamped into `1..4`.
     * @param block extra run formatting layered on top of the bold weight and size.
     * @return the created [XWPFParagraph].
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

    /**
     * Starts a new page by adding an empty paragraph flagged as a page break,
     * so that following content begins on the next page.
     */
    public fun pageBreak() {
        poiDocument.createParagraph().isPageBreak = true
    }

    /**
     * Adds a table and populates it via [block].
     *
     * Rows are added with `row { }` and cells with `cell("...")`. A freshly
     * created POI table already has one empty row and cell, which the first
     * `row`/`cell` calls fill before new ones are appended.
     *
     * ```kotlin
     * table(width = "100%") {
     *     row {
     *         cell("Metric")
     *         cell("Value")
     *     }
     *     row {
     *         cell("Revenue")
     *         cell("12M")
     *     }
     * }
     * ```
     *
     * @param width optional table width: a percentage like `"100%"`, `"auto"`,
     *   or a twips value like `"5000"`; when null POI's default width is used.
     * @param block configures the table's rows and cells.
     * @return the created [XWPFTable].
     */
    public fun table(width: String? = null, block: TableBuilder.() -> Unit): XWPFTable {
        val poiTable = poiDocument.createTable()
        width?.let { poiTable.setWidth(it) }
        TableBuilder(poiTable).block()
        return poiTable
    }
}
