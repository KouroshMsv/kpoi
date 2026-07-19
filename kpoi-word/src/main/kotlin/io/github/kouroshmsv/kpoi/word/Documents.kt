package io.github.kouroshmsv.kpoi.word

import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/**
 * Builds a Word document in memory from a [DocumentBuilder] block.
 *
 * The returned [XWPFDocument] is `AutoCloseable`; close it (or use the [docx]
 * one-liner) once you have written it out, so POI can release its temporary
 * files. The [writeTo] helpers and [toByteArray] do not close it for you.
 *
 * ```kotlin
 * val bytes = document {
 *     heading("Report", level = 1)
 *     paragraph("Hello, world")
 * }.use { it.toByteArray() }
 * ```
 *
 * @param block builder lambda invoked against the new document.
 * @return the populated, still-open [XWPFDocument].
 */
public fun document(block: DocumentBuilder.() -> Unit): XWPFDocument =
    XWPFDocument().also { DocumentBuilder(it).block() }

/**
 * Builds a document with [block] and writes it to a `.docx` file at [path],
 * then closes it so all resources are released.
 *
 * ```kotlin
 * docx(Path.of("letter.docx")) {
 *     heading("Quarterly Letter", level = 1)
 *     paragraph("Dear shareholders,")
 * }
 * ```
 *
 * @param path destination file; created if absent, truncated if present.
 * @param block builder lambda invoked against the new document.
 */
public fun docx(path: Path, block: DocumentBuilder.() -> Unit) {
    document(block).use { it.writeTo(path) }
}

/**
 * Writes the document in `.docx` (OOXML) form to [out].
 *
 * @param out destination stream; left open for the caller to close.
 */
public fun XWPFDocument.writeTo(out: OutputStream) {
    write(out)
}

/**
 * Writes the document to a `.docx` file at [path]. The output stream is
 * opened and closed by this call.
 *
 * @param path destination file; created if absent, truncated if present.
 */
public fun XWPFDocument.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/**
 * Writes the document to a `.docx` [file].
 *
 * @param file destination file; created if absent, truncated if present.
 */
public fun XWPFDocument.writeTo(file: File) {
    writeTo(file.toPath())
}

/**
 * Renders the document to a `.docx` byte array, e.g. for an HTTP response
 * body. The document itself is not closed.
 *
 * @return the complete `.docx` file contents.
 */
public fun XWPFDocument.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
