package dev.kpoi.word

import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/** Builds a Word document in memory. Close it (or use [docx]) when done. */
public fun document(block: DocumentBuilder.() -> Unit): XWPFDocument =
    XWPFDocument().also { DocumentBuilder(it).block() }

/** Builds a DOCX file at [path] and releases all resources. */
public fun docx(path: Path, block: DocumentBuilder.() -> Unit) {
    document(block).use { it.writeTo(path) }
}

/** Writes the document to [out]. The stream is not closed. */
public fun XWPFDocument.writeTo(out: OutputStream) {
    write(out)
}

/** Writes the document to [path]. */
public fun XWPFDocument.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/** Writes the document to [file]. */
public fun XWPFDocument.writeTo(file: File) {
    writeTo(file.toPath())
}

/** Renders the document to a byte array, e.g. for an HTTP response. */
public fun XWPFDocument.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
