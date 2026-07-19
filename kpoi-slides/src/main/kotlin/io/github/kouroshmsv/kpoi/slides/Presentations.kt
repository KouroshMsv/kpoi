package io.github.kouroshmsv.kpoi.slides

import org.apache.poi.xslf.usermodel.XMLSlideShow
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Path

/**
 * Builds a PowerPoint presentation in memory and returns the underlying
 * [XMLSlideShow]. The caller owns the result: close it when done (or use
 * [pptx], which closes it for you) so POI can release its resources.
 *
 * Configure the deck through the [PresentationBuilder] receiver — set the page
 * size and add slides — then persist it with [writeTo] or [toByteArray].
 *
 * ```kotlin
 * val show = presentation {
 *     widescreen()
 *     slide(SlideLayout.TITLE_ONLY) {
 *         title("Quarterly Update")
 *     }
 * }
 * show.use { it.writeTo(Path.of("deck.pptx")) }
 * ```
 *
 * @param block configures the deck through the [PresentationBuilder] receiver.
 * @return an open [XMLSlideShow]; the caller is responsible for closing it.
 */
public fun presentation(block: PresentationBuilder.() -> Unit): XMLSlideShow =
    XMLSlideShow().also { PresentationBuilder(it).block() }

/**
 * Builds a presentation and writes it to a `.pptx` file at [path], then closes
 * the underlying [XMLSlideShow] and releases all POI resources before
 * returning. An existing file is overwritten.
 *
 * Prefer this over [presentation] when you just want a file on disk and do not
 * need to keep the slide show open.
 *
 * ```kotlin
 * pptx(Path.of("deck.pptx")) {
 *     widescreen()
 *     slide(SlideLayout.TITLE_AND_CONTENT) {
 *         title("Next quarter")
 *         placeholder(1) {
 *             bullets("Ship v1.0", "Grow the community")
 *         }
 *     }
 * }
 * ```
 *
 * @param path destination file; created if absent, overwritten if present.
 * @param block configures the deck through the [PresentationBuilder] receiver.
 */
public fun pptx(path: Path, block: PresentationBuilder.() -> Unit) {
    presentation(block).use { it.writeTo(path) }
}

/**
 * Writes the presentation to [out] in `.pptx` (OOXML) form. The stream is left
 * open, so the caller retains ownership and is responsible for closing it.
 *
 * @param out destination stream; not closed by this call.
 */
public fun XMLSlideShow.writeTo(out: OutputStream) {
    write(out)
}

/**
 * Writes the presentation to a `.pptx` file at [path], creating it or
 * truncating an existing file. The output stream it opens is closed for you.
 *
 * @param path destination file.
 */
public fun XMLSlideShow.writeTo(path: Path) {
    Files.newOutputStream(path).use { write(it) }
}

/**
 * Writes the presentation to a `.pptx` [file], creating it or truncating an
 * existing one. Convenience overload of [writeTo] that accepts a [File].
 *
 * @param file destination file.
 */
public fun XMLSlideShow.writeTo(file: File) {
    writeTo(file.toPath())
}

/**
 * Serializes the presentation to a `.pptx` byte array — handy for returning a
 * deck directly in an HTTP response instead of writing it to disk.
 *
 * @return the complete `.pptx` document as a byte array.
 */
public fun XMLSlideShow.toByteArray(): ByteArray =
    ByteArrayOutputStream(64 * 1024).also { write(it) }.toByteArray()
