package io.github.kouroshmsv.kpoi.word

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow

/**
 * DSL scope for one table, and the receiver of [DocumentBuilder.table].
 *
 * Add rows with [row]. The wrapped POI table is exposed as [poiTable].
 */
@PoiDsl
public class TableBuilder internal constructor(
    /** The underlying POI table, for anything the DSL does not cover. */
    public val poiTable: XWPFTable,
) {
    private var nextRowIndex: Int = 0

    /**
     * Adds the next table row, or fills the pre-existing one. A freshly created
     * POI table already has a first row, so the first `row { }` call fills it
     * and later calls append new rows.
     *
     * @param block configures the row's cells.
     * @return the [XWPFTableRow] that was filled or created.
     */
    public fun row(block: TableRowBuilder.() -> Unit = {}): XWPFTableRow {
        val poiRow = if (nextRowIndex < poiTable.numberOfRows) {
            poiTable.getRow(nextRowIndex)
        } else {
            poiTable.createRow()
        }
        nextRowIndex++
        TableRowBuilder(poiRow).block()
        return poiRow
    }
}

/**
 * DSL scope for one table row, and the receiver of [TableBuilder.row].
 *
 * Add cells with [cell]. The wrapped POI row is exposed as [poiRow].
 */
@PoiDsl
public class TableRowBuilder internal constructor(
    /** The underlying POI table row, for anything the DSL does not cover. */
    public val poiRow: XWPFTableRow,
) {
    private var nextCellIndex: Int = 0

    /**
     * Adds the next cell, or fills the pre-existing one, and returns it.
     *
     * When [text] is given it becomes the first paragraph's text; [block] then
     * runs against that first paragraph as a [ParagraphBuilder], so you can
     * align it or add styled runs.
     *
     * ```kotlin
     * row {
     *     cell("Metric")
     *     cell {
     *         text("Value") { bold = true }
     *     }
     * }
     * ```
     *
     * @param text optional text for the cell's first paragraph.
     * @param block formats that first paragraph.
     * @return the [XWPFTableCell] that was filled or created.
     */
    public fun cell(text: String? = null, block: ParagraphBuilder.() -> Unit = {}): XWPFTableCell {
        val poiCell = poiRow.getCell(nextCellIndex) ?: poiRow.addNewTableCell()
        nextCellIndex++
        val poiParagraph = poiCell.paragraphs.firstOrNull() ?: poiCell.addParagraph()
        val builder = ParagraphBuilder(poiParagraph)
        if (text != null) {
            builder.text(text)
        }
        builder.block()
        builder.applyProperties()
        return poiCell
    }
}
