package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.Date

/**
 * DSL scope for a single row; the receiver of the [SheetBuilder.row] and
 * [SheetBuilder.rowAt] lambdas.
 *
 * Cells are appended left to right by the [cell] overloads (one per value type),
 * plus [formula] and [hyperlink]; [skipCells] leaves gaps and [height] sets the
 * row height. A style passed to the row becomes each cell's default and merges
 * with per-cell styles.
 */
@PoiDsl
public class RowBuilder internal constructor(
    private val workbook: WorkbookBuilder,
    /** Escape hatch to the underlying POI [Row] for anything the DSL does not cover. */
    public val poiRow: Row,
    private val rowStyle: StyleSpec,
) {
    private var nextColumnIndex: Int = 0

    /**
     * Sets the row height in points.
     *
     * @param points height in typographic points.
     */
    public fun height(points: Float) {
        poiRow.heightInPoints = points
    }

    /**
     * Leaves [count] cells empty before the next cell is written.
     *
     * @param count number of columns to skip.
     */
    public fun skipCells(count: Int) {
        nextColumnIndex += count
    }

    /**
     * Adds a text cell. A `null` [value] creates a blank cell that still carries
     * the row, handle, and inline style, which is handy for keeping borders or
     * fills in otherwise empty positions.
     *
     * ```kotlin
     * row {
     *     cell("Alice")
     *     cell(null)                                    // styled blank cell
     *     cell("Total") { style { font { bold = true } } }
     * }
     * ```
     *
     * @param value text to write, or `null` for a styled blank cell.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration: inline [CellBuilder.style],
     *   [CellBuilder.format], or [CellBuilder.hyperlink].
     * @return the created POI [Cell].
     */
    public fun cell(value: String?, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { if (value != null) it.setCellValue(value) }

    /**
     * Adds a numeric cell. The [value] is stored as a `Double` (POI's only
     * numeric type), so any [Number] — `Int`, `Long`, `BigDecimal`, etc. — works.
     *
     * @param value the number to store.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun cell(value: Number, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.setCellValue(value.toDouble()) }

    /**
     * Adds a boolean cell (Excel `TRUE`/`FALSE`).
     *
     * @param value the boolean to store.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun cell(value: Boolean, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.setCellValue(value) }

    /**
     * Adds a date cell. Unless a format is set on the cell or its style, the
     * workbook's [WorkbookBuilder.defaultDateFormat] (`"yyyy-mm-dd"`) is applied.
     *
     * @param value the date to store.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun cell(value: LocalDate, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateFormat, block) { it.setCellValue(value) }

    /**
     * Adds a date-time cell. Unless a format is set on the cell or its style, the
     * workbook's [WorkbookBuilder.defaultDateTimeFormat] (`"yyyy-mm-dd hh:mm"`)
     * is applied.
     *
     * @param value the date-time to store.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun cell(value: LocalDateTime, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateTimeFormat, block) { it.setCellValue(value) }

    /**
     * Adds a cell from a legacy [Date], stored as a date-time. Unless a format is
     * set on the cell or its style, the workbook's
     * [WorkbookBuilder.defaultDateTimeFormat] (`"yyyy-mm-dd hh:mm"`) is applied.
     *
     * @param value the date to store.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun cell(value: Date, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, workbook.defaultDateTimeFormat, block) { it.setCellValue(value) }

    /**
     * Adds a formula cell. Give [expression] without the leading `=`, e.g.
     * `"SUM(B2:B9)"`. POI stores the formula and Excel computes its value on open.
     *
     * @param expression the formula text, without a leading `=`.
     * @param style optional style handle merged over the row style.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
    public fun formula(expression: String, style: CellStyleHandle? = null, block: CellBuilder.() -> Unit = {}): Cell =
        addCell(style, null, block) { it.cellFormula = expression }

    /**
     * Adds a clickable hyperlink cell. Unless [style] is given, the cell is
     * styled blue and underlined like a conventional link.
     *
     * @param url the link target.
     * @param label visible text; defaults to [url].
     * @param style optional style handle replacing the default blue-underline look.
     * @param block per-cell configuration.
     * @return the created POI [Cell].
     */
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

/**
 * Per-cell configuration applied inside a `cell(...) { }` block. Lets a single
 * cell override its number [format], add an inline [style], or become a
 * [hyperlink], on top of the row default and any style handle.
 */
@PoiDsl
public class CellBuilder internal constructor(
    /** Escape hatch to the underlying POI [Cell] for anything the DSL does not cover. */
    public val poiCell: Cell,
) {
    /**
     * Number format for just this cell, e.g. `"#,##0.00"`. Takes precedence over
     * any format coming from the row style, a style handle, or a default date
     * format.
     */
    public var format: String? = null

    internal var inlineStyle: StyleSpec? = null
    internal var hyperlinkUrl: String? = null

    /**
     * One-off style for this cell, merged over the row and handle styles (this
     * inline style wins on conflicts). Identical inline styles are still
     * deduplicated workbook-wide, so styling every cell inline stays cheap.
     *
     * @param block style configuration for this cell only.
     */
    public fun style(block: StyleBuilder.() -> Unit) {
        val built = StyleBuilder().apply(block).build()
        inlineStyle = inlineStyle?.mergedWith(built) ?: built
    }

    /**
     * Turns this cell into a hyperlink to [url]. Unlike [RowBuilder.hyperlink]
     * this does not apply the default blue-underline link styling.
     *
     * @param url the link target.
     */
    public fun hyperlink(url: String) {
        hyperlinkUrl = url
    }
}
