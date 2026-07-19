package io.github.kouroshmsv.kpoi.spreadsheet

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/**
 * Builds an in-memory XLSX ([XSSFWorkbook]) from the DSL in [block].
 *
 * The workbook is fully materialized in memory. You own the returned instance:
 * write it out and then close it to release resources, or use [xlsx] which does
 * both for you. For exports too large to hold in memory use [streamingWorkbook];
 * for the legacy `.xls` binary format use [xlsWorkbook].
 *
 * ```kotlin
 * val wb: XSSFWorkbook = workbook {
 *     val header = style { font { bold = true } }
 *     sheet("Report") {
 *         row(header) {
 *             cell("Name")
 *             cell("Score")
 *         }
 *         row {
 *             cell("Alice")
 *             cell(95.5) { format = "0.00" }
 *         }
 *     }
 * }
 * wb.use { it.writeTo(Path.of("report.xlsx")) }
 * ```
 *
 * @param block DSL that configures the workbook: styles, sheets, rows, and cells.
 * @return an open [XSSFWorkbook]; the caller must close it (directly or via `use`).
 */
public fun workbook(block: WorkbookBuilder.() -> Unit): XSSFWorkbook =
    XSSFWorkbook().also { WorkbookBuilder(it).block() }

/**
 * Builds an in-memory legacy XLS ([HSSFWorkbook]) from the DSL in [block].
 *
 * Prefer [workbook] (XLSX) unless you specifically need the old `.xls` binary
 * format. XLS is limited to 65 536 rows, 256 columns, and 4 000 cell styles, and
 * it does not support hex colors: use the indexed-color overloads of
 * [StyleBuilder.fill] and [FontBuilder.color] instead. Close the workbook (or use
 * [xls]) when done.
 *
 * @param block DSL that configures the workbook.
 * @return an open [HSSFWorkbook]; the caller must close it.
 */
public fun xlsWorkbook(block: WorkbookBuilder.() -> Unit): HSSFWorkbook =
    HSSFWorkbook().also { WorkbookBuilder(it).block() }

/**
 * Builds a streaming XLSX ([SXSSFWorkbook]) that flushes rows to a temporary
 * file, keeping at most [rowAccessWindowSize] of the most recently written rows
 * in memory. Use this for very large exports that would otherwise exhaust the
 * heap.
 *
 * Because older rows are flushed to disk they can no longer be revisited, which
 * changes a few features: to use [SheetBuilder.autoSizeColumns] you must call
 * [SheetBuilder.trackColumnsForAutoSize] before writing any rows. Closing the
 * workbook also deletes its temporary files; `xlsx(path, streaming = true)`
 * creates, writes, and cleans up for you.
 *
 * @param rowAccessWindowSize number of rows kept in memory; defaults to
 *   [SXSSFWorkbook.DEFAULT_WINDOW_SIZE]. A larger window trades memory for the
 *   ability to keep more rows editable.
 * @param block DSL that configures the workbook.
 * @return an open [SXSSFWorkbook]; the caller must close it to flush and clean up.
 */
public fun streamingWorkbook(
    rowAccessWindowSize: Int = SXSSFWorkbook.DEFAULT_WINDOW_SIZE,
    block: WorkbookBuilder.() -> Unit,
): SXSSFWorkbook =
    SXSSFWorkbook(rowAccessWindowSize).also { WorkbookBuilder(it).block() }

/**
 * Builds an XLSX file at [path] and releases all resources when done, including
 * any SXSSF temporary files. This is the one-liner form of [workbook] (or
 * [streamingWorkbook] when [streaming] is `true`) followed by [writeTo] and a
 * `close`.
 *
 * ```kotlin
 * xlsx(Path.of("report.xlsx")) {
 *     sheet("Report") {
 *         row {
 *             cell("Name")
 *             cell("Score")
 *         }
 *         row {
 *             cell("Alice")
 *             cell(95.5) { format = "0.00" }
 *         }
 *     }
 * }
 * ```
 *
 * @param path destination file; it is created or overwritten.
 * @param streaming when `true`, uses a memory-friendly [streamingWorkbook] with
 *   the default window size instead of a fully in-memory [workbook].
 * @param block DSL that configures the workbook.
 */
public fun xlsx(path: Path, streaming: Boolean = false, block: WorkbookBuilder.() -> Unit) {
    val built: Workbook = if (streaming) streamingWorkbook(block = block) else workbook(block)
    built.use { it.writeTo(path) }
}

/**
 * Builds a legacy XLS file at [path] and releases all resources. The one-liner
 * form of [xlsWorkbook] followed by [writeTo] and a `close`. Hex colors are not
 * supported in this format; use the indexed-color overloads of [StyleBuilder.fill]
 * and [FontBuilder.color] instead.
 *
 * @param path destination file; it is created or overwritten.
 * @param block DSL that configures the workbook.
 */
public fun xls(path: Path, block: WorkbookBuilder.() -> Unit) {
    xlsWorkbook(block).use { it.writeTo(path) }
}

/**
 * Writes the workbook to [out] in its native format (XLSX or XLS). POI flushes
 * the stream but does not close it, so the caller remains responsible for
 * closing [out].
 *
 * @param out destination stream, left open.
 */
public fun Workbook.writeTo(out: OutputStream) {
    write(out)
}

/**
 * Writes the workbook to [path], creating the file or overwriting it if present.
 *
 * @param path destination file.
 */
public fun Workbook.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/**
 * Writes the workbook to [file], creating it or overwriting it if present.
 *
 * @param file destination file.
 */
public fun Workbook.writeTo(file: File) {
    writeTo(file.toPath())
}

/**
 * Renders the workbook to a byte array in its native format, e.g. to serve as an
 * HTTP response body or persist as a blob. Does not close the workbook.
 *
 * @return the encoded `.xlsx` or `.xls` bytes.
 */
public fun Workbook.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
