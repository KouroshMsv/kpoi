package dev.kpoi.spreadsheet

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/** Builds an XLSX workbook in memory. Close it (or use [xlsx]) when done. */
public fun workbook(block: WorkbookBuilder.() -> Unit): XSSFWorkbook =
    XSSFWorkbook().also { WorkbookBuilder(it).block() }

/** Builds a legacy XLS workbook in memory. Close it (or use [xls]) when done. */
public fun xlsWorkbook(block: WorkbookBuilder.() -> Unit): HSSFWorkbook =
    HSSFWorkbook().also { WorkbookBuilder(it).block() }

/**
 * Builds a streaming XLSX workbook that keeps only [rowAccessWindowSize] rows
 * in memory, for very large exports. Closing it also deletes its temporary
 * files; `xlsx(path, streaming = true)` handles all of that for you.
 */
public fun streamingWorkbook(
    rowAccessWindowSize: Int = SXSSFWorkbook.DEFAULT_WINDOW_SIZE,
    block: WorkbookBuilder.() -> Unit,
): SXSSFWorkbook =
    SXSSFWorkbook(rowAccessWindowSize).also { WorkbookBuilder(it).block() }

/** Builds an XLSX file at [path] and releases all resources (including SXSSF temp files). */
public fun xlsx(path: Path, streaming: Boolean = false, block: WorkbookBuilder.() -> Unit) {
    val built: Workbook = if (streaming) streamingWorkbook(block = block) else workbook(block)
    built.use { it.writeTo(path) }
}

/** Builds a legacy XLS file at [path] and releases all resources. */
public fun xls(path: Path, block: WorkbookBuilder.() -> Unit) {
    xlsWorkbook(block).use { it.writeTo(path) }
}

/** Writes the workbook to [out]. The stream is not closed. */
public fun Workbook.writeTo(out: OutputStream) {
    write(out)
}

/** Writes the workbook to [path]. */
public fun Workbook.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/** Writes the workbook to [file]. */
public fun Workbook.writeTo(file: File) {
    writeTo(file.toPath())
}

/** Renders the workbook to a byte array, e.g. for an HTTP response. */
public fun Workbook.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
