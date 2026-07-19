package dev.kpoi.spreadsheet

import dev.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.WorkbookUtil

/** Top-level DSL scope for building a workbook. */
@PoiDsl
public class WorkbookBuilder internal constructor(
    /** The underlying POI workbook, for anything the DSL does not cover. */
    public val poiWorkbook: Workbook,
) {
    internal val styleRegistry: StyleRegistry = StyleRegistry(poiWorkbook)

    /** Number format applied to date cells that have no explicit format. */
    public var defaultDateFormat: String = "yyyy-mm-dd"

    /** Number format applied to date-time cells that have no explicit format. */
    public var defaultDateTimeFormat: String = "yyyy-mm-dd hh:mm"

    /**
     * Defines a reusable style. Handles are cheap descriptions; the POI style
     * is created lazily and shared by every cell with an identical look.
     */
    public fun style(block: StyleBuilder.() -> Unit): CellStyleHandle =
        CellStyleHandle(StyleBuilder().apply(block).build())

    /** Adds a sheet. [name] is sanitized with [WorkbookUtil.createSafeSheetName]. */
    public fun sheet(name: String? = null, block: SheetBuilder.() -> Unit = {}): Sheet {
        val poiSheet = if (name == null) {
            poiWorkbook.createSheet()
        } else {
            poiWorkbook.createSheet(WorkbookUtil.createSafeSheetName(name))
        }
        SheetBuilder(this, poiSheet).block()
        return poiSheet
    }
}
