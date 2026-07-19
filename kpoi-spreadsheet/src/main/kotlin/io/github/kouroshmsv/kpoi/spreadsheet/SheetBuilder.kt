package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFSheet

/** DSL scope for one sheet. */
@PoiDsl
public class SheetBuilder internal constructor(
    private val workbook: WorkbookBuilder,
    /** The underlying POI sheet, for anything the DSL does not cover. */
    public val poiSheet: Sheet,
) {
    private var nextRowIndex: Int = 0

    /** Appends a row. [style] becomes the default for all its cells. */
    public fun row(style: CellStyleHandle? = null, block: RowBuilder.() -> Unit = {}): Row =
        rowAt(nextRowIndex, style, block)

    /** Creates a row at an explicit zero-based [index] and continues from there. */
    public fun rowAt(index: Int, style: CellStyleHandle? = null, block: RowBuilder.() -> Unit = {}): Row {
        val poiRow = poiSheet.createRow(index)
        nextRowIndex = index + 1
        RowBuilder(workbook, poiRow, style?.spec ?: StyleSpec.EMPTY).block()
        return poiRow
    }

    /** Leaves [count] empty rows before the next [row] call. */
    public fun skipRows(count: Int) {
        nextRowIndex += count
    }

    /** Sets a column width in characters, as shown in Excel's resize tooltip. */
    public fun columnWidth(column: Int, characters: Int) {
        poiSheet.setColumnWidth(column, characters * 256)
    }

    /**
     * Auto-sizes the given columns to fit their content; call after the data
     * is written. On streaming (SXSSF) sheets, call [trackColumnsForAutoSize]
     * before writing any rows or POI will throw.
     */
    public fun autoSizeColumns(vararg columns: Int) {
        columns.forEach { poiSheet.autoSizeColumn(it) }
    }

    /** Enables column tracking on streaming sheets so [autoSizeColumns] works. No-op otherwise. */
    public fun trackColumnsForAutoSize() {
        (poiSheet as? SXSSFSheet)?.trackAllColumnsForAutoSizing()
    }

    /** Freezes the top [rows] and/or left [columns]. */
    public fun freeze(rows: Int = 0, columns: Int = 0) {
        poiSheet.createFreezePane(columns, rows)
    }

    /** Merges a range given in A1 notation, e.g. `"A1:C1"`. */
    public fun merge(range: String) {
        poiSheet.addMergedRegion(CellRangeAddress.valueOf(range))
    }

    /** Merges a range given as zero-based row/column bounds (inclusive). */
    public fun merge(firstRow: Int, lastRow: Int, firstColumn: Int, lastColumn: Int) {
        poiSheet.addMergedRegion(CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn))
    }

    /** Adds an auto-filter dropdown over the given A1 range, e.g. `"A1:D1"`. */
    public fun autoFilter(range: String) {
        poiSheet.setAutoFilter(CellRangeAddress.valueOf(range))
    }
}
