package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.WorkbookUtil

/**
 * Top-level DSL scope for building a workbook; the receiver of the [workbook],
 * [xlsWorkbook], [streamingWorkbook], [xlsx], and [xls] lambdas.
 *
 * From here you declare reusable styles with [style] and add sheets with [sheet].
 * The same builder serves XLSX, streaming XLSX, and legacy XLS — the concrete
 * format is fixed by the entry point that created it. Styles are content-keyed
 * and deduplicated across the whole workbook (see [style]).
 */
@PoiDsl
public class WorkbookBuilder internal constructor(
    /** Escape hatch to the underlying POI [Workbook] for anything the DSL does not cover. */
    public val poiWorkbook: Workbook,
) {
    internal val styleRegistry: StyleRegistry = StyleRegistry(poiWorkbook)

    /**
     * Excel number format applied to [java.time.LocalDate] cells that set no
     * explicit format of their own. Assign a different pattern to change it
     * workbook-wide.
     */
    public var defaultDateFormat: String = "yyyy-mm-dd"

    /**
     * Excel number format applied to [java.time.LocalDateTime] and
     * [java.util.Date] cells that set no explicit format of their own.
     */
    public var defaultDateTimeFormat: String = "yyyy-mm-dd hh:mm"

    /**
     * Defines a reusable [CellStyleHandle]. A handle is a cheap, immutable
     * description of a look; the actual POI `CellStyle` is created lazily and
     * shared by every cell with an identical style, so declaring one handle and
     * applying it to a million cells still yields a single style in the file.
     *
     * The handle can be passed to [SheetBuilder.row] (as a row default) or to any
     * `cell(...)` overload, and it merges with row and inline styles (see
     * [CellBuilder.style]).
     *
     * ```kotlin
     * val header = style {
     *     font {
     *         bold = true
     *         color("#FFFFFF")
     *     }
     *     fill("#4472C4")
     *     align(horizontal = HorizontalAlignment.CENTER)
     * }
     * sheet("Report") {
     *     row(header) { cell("Name") }
     * }
     * ```
     *
     * @param block style configuration: font, fill, alignment, borders, format.
     * @return a reusable handle you can apply to rows and cells.
     */
    public fun style(block: StyleBuilder.() -> Unit): CellStyleHandle =
        CellStyleHandle(StyleBuilder().apply(block).build())

    /**
     * Adds a sheet and configures it with [block].
     *
     * ```kotlin
     * sheet("Report") {
     *     freeze(rows = 1)
     *     columnWidth(0, characters = 18)
     *     row {
     *         cell("Name")
     *         cell("Score")
     *     }
     * }
     * ```
     *
     * @param name sheet name; when non-`null` it is sanitized with
     *   [WorkbookUtil.createSafeSheetName] (invalid characters replaced, trimmed
     *   to 31 characters). When `null`, POI assigns a default name.
     * @param block DSL that adds rows and configures sheet-level features.
     * @return the created POI [Sheet].
     */
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
