package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFSheet

/**
 * DSL scope for a single sheet; the receiver of the [WorkbookBuilder.sheet]
 * lambda.
 *
 * Rows are normally appended in order with [row]; [rowAt] jumps to an explicit
 * index and [skipRows] leaves gaps. Sheet-level features include [columnWidth],
 * [freeze], [merge], [autoFilter], and [autoSizeColumns].
 */
@PoiDsl
public class SheetBuilder internal constructor(
    private val workbook: WorkbookBuilder,
    /** Escape hatch to the underlying POI [Sheet] for anything the DSL does not cover. */
    public val poiSheet: Sheet,
) {
    private var nextRowIndex: Int = 0

    /**
     * Appends a row after the last one created and fills it with [block].
     *
     * ```kotlin
     * row {
     *     cell("Alice")
     *     cell(95.5) { format = "0.00" }
     * }
     * ```
     *
     * @param style optional default style for every cell in the row; individual
     *   cells can still override or extend it (styles merge, see [CellBuilder.style]).
     * @param block DSL that adds cells to the row.
     * @return the created POI [Row].
     */
    public fun row(style: CellStyleHandle? = null, block: RowBuilder.() -> Unit = {}): Row =
        rowAt(nextRowIndex, style, block)

    /**
     * Creates a row at the explicit zero-based [index], replacing any existing
     * row there, and resumes sequential [row] appends from `index + 1`.
     *
     * @param index zero-based row index.
     * @param style optional default style for the row's cells.
     * @param block DSL that adds cells to the row.
     * @return the created POI [Row].
     */
    public fun rowAt(index: Int, style: CellStyleHandle? = null, block: RowBuilder.() -> Unit = {}): Row {
        val poiRow = poiSheet.createRow(index)
        nextRowIndex = index + 1
        RowBuilder(workbook, poiRow, style?.spec ?: StyleSpec.EMPTY).block()
        return poiRow
    }

    /**
     * Leaves [count] empty rows before the next [row] call. Has no effect on
     * [rowAt], which sets the position explicitly.
     *
     * @param count number of rows to skip.
     */
    public fun skipRows(count: Int) {
        nextRowIndex += count
    }

    /**
     * Sets the width of [column] to [characters] character widths, matching the
     * number Excel shows in its column-resize tooltip.
     *
     * @param column zero-based column index.
     * @param characters approximate width in characters.
     */
    public fun columnWidth(column: Int, characters: Int) {
        poiSheet.setColumnWidth(column, characters * 256)
    }

    /**
     * Auto-sizes the given [columns] to fit their content; call after the data
     * is written. On streaming (SXSSF) sheets you must call
     * [trackColumnsForAutoSize] before writing any rows, or POI throws when it
     * tries to measure rows already flushed to disk.
     *
     * @param columns zero-based indices of the columns to size.
     */
    public fun autoSizeColumns(vararg columns: Int) {
        columns.forEach { poiSheet.autoSizeColumn(it) }
    }

    /**
     * Enables column tracking on streaming (SXSSF) sheets so [autoSizeColumns]
     * can measure content later; call it before writing any rows. On regular
     * XLSX/XLS sheets widths are always measurable, so this is a no-op.
     */
    public fun trackColumnsForAutoSize() {
        (poiSheet as? SXSSFSheet)?.trackAllColumnsForAutoSizing()
    }

    /**
     * Freezes the top [rows] and/or left [columns] so they stay visible while
     * scrolling; `freeze(rows = 1)` pins a header row.
     *
     * @param rows number of rows to freeze at the top.
     * @param columns number of columns to freeze on the left.
     */
    public fun freeze(rows: Int = 0, columns: Int = 0) {
        poiSheet.createFreezePane(columns, rows)
    }

    /**
     * Merges the cells in an A1-notation [range], e.g. `"A1:C1"`. The value and
     * style of the top-left cell apply to the whole region.
     *
     * @param range merged region in A1 notation.
     */
    public fun merge(range: String) {
        poiSheet.addMergedRegion(CellRangeAddress.valueOf(range))
    }

    /**
     * Merges a region given as zero-based, inclusive row and column bounds.
     *
     * @param firstRow top row (inclusive).
     * @param lastRow bottom row (inclusive).
     * @param firstColumn left column (inclusive).
     * @param lastColumn right column (inclusive).
     */
    public fun merge(firstRow: Int, lastRow: Int, firstColumn: Int, lastColumn: Int) {
        poiSheet.addMergedRegion(CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn))
    }

    /**
     * Adds an auto-filter dropdown over the given A1 [range], typically the
     * header row, e.g. `"A1:D1"`.
     *
     * @param range filtered range in A1 notation.
     */
    public fun autoFilter(range: String) {
        poiSheet.setAutoFilter(CellRangeAddress.valueOf(range))
    }
}
