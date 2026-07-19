package io.github.kouroshmsv.kpoi.word

import io.github.kouroshmsv.kpoi.common.PoiDsl
import io.github.kouroshmsv.kpoi.common.Rgb
import org.apache.poi.common.usermodel.PictureType
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import java.nio.file.Files
import java.nio.file.Path

/**
 * DSL scope for one paragraph, and the receiver of [DocumentBuilder.paragraph]
 * and [TableRowBuilder.cell].
 *
 * Add runs with [text], blank line breaks with [lineBreak], and images with
 * [picture]; set paragraph-level [alignment] and spacing via
 * [spacingBeforePoints] / [spacingAfterPoints]. The wrapped POI paragraph is
 * exposed as [poiParagraph].
 */
@PoiDsl
public class ParagraphBuilder internal constructor(
    /** The underlying POI paragraph, for anything the DSL does not cover. */
    public val poiParagraph: XWPFParagraph,
) {
    /** Horizontal alignment; null leaves the paragraph at Word's default (left). */
    public var alignment: ParagraphAlignment? = null

    /** Space before the paragraph, in points; null leaves POI's default. */
    public var spacingBeforePoints: Int? = null

    /** Space after the paragraph, in points; null leaves POI's default. */
    public var spacingAfterPoints: Int? = null

    /**
     * Adds a run of [value] to the paragraph. Any embedded `\n` characters
     * become line breaks (`<w:br/>`) within this single run, so the whole
     * string stays one styled run. Character formatting is applied via [block].
     *
     * ```kotlin
     * text("first line\nsecond line") {
     *     bold = true
     *     color = "#C00000"
     * }
     * ```
     *
     * @param value run text; each `\n` becomes a line break.
     * @param block character formatting for the run.
     * @return the created [XWPFRun].
     */
    public fun text(value: String, block: RunBuilder.() -> Unit = {}): XWPFRun {
        val run = poiParagraph.createRun()
        value.split('\n').forEachIndexed { index, line ->
            if (index > 0) {
                run.addBreak()
            }
            run.setText(line, index)
        }
        RunBuilder().apply(block).applyTo(run)
        return run
    }

    /** Adds a line break (`<w:br/>`) as its own, otherwise empty, run. */
    public fun lineBreak() {
        poiParagraph.createRun().addBreak()
    }

    /**
     * Embeds the image at [path] as a new run, scaled to [widthPoints] x
     * [heightPoints] points (converted to EMUs for POI). The picture type is
     * inferred from the file extension.
     *
     * Supported extensions are `png`, `jpg`, `jpeg`, `gif`, and `bmp` (matched
     * case-insensitively); any other extension throws.
     *
     * ```kotlin
     * paragraph {
     *     picture(Path.of("logo.png"), widthPoints = 120, heightPoints = 40)
     * }
     * ```
     *
     * @param path image file to read and embed.
     * @param widthPoints rendered width in points.
     * @param heightPoints rendered height in points.
     * @return the created [XWPFRun] holding the picture.
     * @throws IllegalArgumentException if the file extension is not one of
     *   `png`, `jpg`, `jpeg`, `gif`, or `bmp`.
     */
    public fun picture(path: Path, widthPoints: Int, heightPoints: Int): XWPFRun {
        val run = poiParagraph.createRun()
        Files.newInputStream(path).use { stream ->
            run.addPicture(
                stream,
                pictureTypeOf(path),
                path.fileName.toString(),
                Units.toEMU(widthPoints.toDouble()),
                Units.toEMU(heightPoints.toDouble()),
            )
        }
        return run
    }

    internal fun applyProperties() {
        alignment?.let { poiParagraph.alignment = it }
        spacingBeforePoints?.let { poiParagraph.spacingBefore = it * TWIPS_PER_POINT }
        spacingAfterPoints?.let { poiParagraph.spacingAfter = it * TWIPS_PER_POINT }
    }

    private companion object {
        private const val TWIPS_PER_POINT = 20

        private fun pictureTypeOf(path: Path): PictureType =
            when (val extension = path.fileName.toString().substringAfterLast('.').lowercase()) {
                "png" -> PictureType.PNG
                "jpg", "jpeg" -> PictureType.JPEG
                "gif" -> PictureType.GIF
                "bmp" -> PictureType.BMP
                else -> throw IllegalArgumentException(
                    "Unsupported image extension \"$extension\"; use png, jpg, jpeg, gif, or bmp"
                )
            }
    }
}

/**
 * Character formatting for one run, applied inside a `text(...) { }` block (and
 * used internally by [DocumentBuilder.heading]).
 *
 * Every property defaults to `null`, which leaves POI's default for that
 * attribute untouched.
 */
@PoiDsl
public class RunBuilder internal constructor() {
    /** Bold weight when true; null leaves the run unchanged. */
    public var bold: Boolean? = null

    /** Italic style when true; null leaves the run unchanged. */
    public var italic: Boolean? = null

    /** Single underline when true, none when false; null leaves the run unchanged. */
    public var underline: Boolean? = null

    /** Strikethrough when true; null leaves the run unchanged. */
    public var strikethrough: Boolean? = null

    /** Font size in points; any [Number], applied as a double. Null leaves the default. */
    public var size: Number? = null

    /**
     * Text color as an `RRGGBB` hex string, with or without a leading `#` and
     * case-insensitive, e.g. `"#C00000"` or `"c00000"`. Null leaves the default.
     */
    public var color: String? = null

    /** Font family name, e.g. `"Calibri"`. Null leaves the default. */
    public var fontFamily: String? = null

    internal fun applyTo(run: XWPFRun) {
        bold?.let { run.isBold = it }
        italic?.let { run.isItalic = it }
        underline?.let { run.underline = if (it) UnderlinePatterns.SINGLE else UnderlinePatterns.NONE }
        strikethrough?.let { run.isStrikeThrough = it }
        size?.let { run.setFontSize(it.toDouble()) }
        color?.let { run.color = Rgb.parse(it).toHexString() }
        fontFamily?.let { run.fontFamily = it }
    }
}
