package dev.kpoi.word

import dev.kpoi.common.PoiDsl
import org.apache.poi.xwpf.usermodel.XWPFTable
import org.apache.poi.xwpf.usermodel.XWPFTableCell
import org.apache.poi.xwpf.usermodel.XWPFTableRow

/** DSL scope for one table. */
@PoiDsl
public class TableBuilder internal constructor(
    /** The underlying POI table, for anything the DSL does not cover. */
    public val poiTable: XWPFTable,
) {
    private var nextRowIndex: Int = 0

    /** Adds (or fills) the next table row. */
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

/** DSL scope for one table row. */
@PoiDsl
public class TableRowBuilder internal constructor(
    /** The underlying POI table row, for anything the DSL does not cover. */
    public val poiRow: XWPFTableRow,
) {
    private var nextCellIndex: Int = 0

    /** Adds (or fills) the next cell; [text] becomes its first paragraph. */
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
