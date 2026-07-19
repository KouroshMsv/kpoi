package io.github.kouroshmsv.kpoi.slides

import org.apache.poi.xslf.usermodel.XMLSlideShow
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/** Builds a PowerPoint presentation in memory. Close it (or use [pptx]) when done. */
public fun presentation(block: PresentationBuilder.() -> Unit): XMLSlideShow =
    XMLSlideShow().also { PresentationBuilder(it).block() }

/** Builds a PPTX file at [path] and releases all resources. */
public fun pptx(path: Path, block: PresentationBuilder.() -> Unit) {
    presentation(block).use { it.writeTo(path) }
}

/** Writes the presentation to [out]. The stream is not closed. */
public fun XMLSlideShow.writeTo(out: OutputStream) {
    write(out)
}

/** Writes the presentation to [path]. */
public fun XMLSlideShow.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/** Writes the presentation to [file]. */
public fun XMLSlideShow.writeTo(file: File) {
    writeTo(file.toPath())
}

/** Renders the presentation to a byte array, e.g. for an HTTP response. */
public fun XMLSlideShow.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
