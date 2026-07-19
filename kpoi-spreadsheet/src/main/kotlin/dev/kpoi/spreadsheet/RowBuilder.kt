package dev.kpoi.spreadsheet

import dev.kpoi.common.PoiDsl
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Date

/** DSL scope for one row. */
@PoiDsl
public class RowBuilder internal constructor(
    private val workbook: WorkbookBuilder,
    /** The underlying POI row, for anything the DSL does not cover. */
    public val poiRow: Row,
    private val rowStyle: StyleSpec,
) {
    private var nextColumnIndex: Int = 0

    /** Sets the row height in points. */
    public fun height(points: Float) {
        poiRow.heightInPoints = points
    }

    /** Leaves [count] cells empty before the next cell. */
    public fun skipCells(count: Int) {
        nextColumnIndex += count
    }

    /** Adds a text cell; `null` produces a styled blank cell. */
    public fun cell(value: String?, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { if (value != null) it.setCellValue(value) }

    /** Adds a numeric cell. */
    public fun cell(value: Number, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.setCellValue(value.toDouble()) }

    /** Adds a boolean cell. */
    public fun cell(value: Boolean, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.setCellValue(value) }

    /** Adds a date cell, applying [WorkbookBuilder.defaultDateFormat] unless overridden. */
    public fun cell(value: LocalDate, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateFormat, block) { it.setCellValue(value) }

    /** Adds a date-time cell, applying [WorkbookBuilder.defaultDateTimeFormat] unless overridden. */
    public fun cell(value: LocalDateTime, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateTimeFormat, block) { it.setCellValue(value) }

    /** Adds a date-time cell, applying [WorkbookBuilder.defaultDateTimeFormat] unless overridden. */
    public fun cell(value: Date, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateTimeFormat, block) { it.setCellValue(value) }

    /** Adds a formula cell; [expression] without the leading `=`, e.g. `"SUM(B2:B9)"`. */
    public fun formula(expression: String, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.cellFormula = expression }

    /** Adds a hyperlink cell, styled blue and underlined unless [style] is given. */
    public fun hyperlink(
        url: String,
        label: String = url,
        style: CellStyleHandle? = null,
        block: CellBuilder.() -> Unit = {},
    ): Cell =
        cell(label, style ?: LINK_STYLE) {
            hyperlink(url)
            block()
        }

    private fun addCell(
        style: CellStyleHandle?,
        defaultFormat: String?,
        block: CellBuilder.() -> Unit,
        write: (Cell) -> Unit,
    ): Cell {
        val poiCell = poiRow.createCell(nextColumnIndex++)
        write(poiCell)
        val cellBuilder = CellBuilder(poiCell).apply(block)
        var spec = rowStyle
        style?.let { spec = spec.mergedWith(it.spec) }
        cellBuilder.inlineStyle?.let { spec = spec.mergedWith(it) }
        cellBuilder.format?.let { spec = spec.copy(dataFormat = it) }
        if (defaultFormat != null && spec.dataFormat == null) {
            spec = spec.copy(dataFormat = defaultFormat)
        }
        if (!spec.isEmpty) {
            poiCell.cellStyle = workbook.styleRegistry.resolve(spec)
        }
        cellBuilder.hyperlinkUrl?.let { linkUrl ->
            val link = workbook.poiWorkbook.creationHelper.createHyperlink(HyperlinkType.URL)
            link.address = linkUrl
            poiCell.hyperlink = link
        }
        return poiCell
    }

    private companion object {
        private val LINK_STYLE = CellStyleHandle(
            StyleSpec(font = FontSpec(underline = true, color = IndexedColors.BLUE)),
        )
    }
}

/** Per-cell configuration applied inside a `cell(...) { }` block. */
@PoiDsl
public class CellBuilder internal constructor(
    /** The underlying POI cell, for anything the DSL does not cover. */
    public val poiCell: Cell,
) {
    /** Number format for just this cell, e.g. `"#,##0.00"`. */
    public var format: String? = null

    internal var inlineStyle: StyleSpec? = null
    internal var hyperlinkUrl: String? = null

    /**
     * One-off style for this cell, merged over the row and handle styles.
     * Identical inline styles are still deduplicated workbook-wide.
     */
    public fun style(block: StyleBuilder.() -> Unit) {
        val built = StyleBuilder().apply(block).build()
        inlineStyle = inlineStyle?.mergedWith(built) ?: built
    }

    /** Turns this cell into a hyperlink to [url]. */
    public fun hyperlink(url: String) {
        hyperlinkUrl = url
    }
}
