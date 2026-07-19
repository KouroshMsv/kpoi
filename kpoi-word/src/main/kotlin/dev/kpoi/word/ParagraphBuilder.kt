package dev.kpoi.word

import dev.kpoi.common.PoiDsl
import dev.kpoi.common.Rgb
import org.apache.poi.common.usermodel.PictureType
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import org.apache.poi.xwpf.usermodel.UnderlinePatterns
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import java.nio.file.Files
import java.nio.file.Path

/** DSL scope for one paragraph. */
@PoiDsl
public class ParagraphBuilder internal constructor(
    /** The underlying POI paragraph, for anything the DSL does not cover. */
    public val poiParagraph: XWPFParagraph,
) {
    public var alignment: ParagraphAlignment? = null

    /** Space before the paragraph, in points. */
    public var spacingBeforePoints: Int? = null

    /** Space after the paragraph, in points. */
    public var spacingAfterPoints: Int? = null

    /** Adds a run of [value]; embedded `\n` become line breaks. */
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

    /** Adds a line break as its own run. */
    public fun lineBreak() {
        poiParagraph.createRun().addBreak()
    }

    /** Embeds an image scaled to the given size in points. */
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
                    "Unsupported image extension \"$extension\"; use png, jpg, gif, or bmp"
                )
            }
    }
}

/** Character formatting for one run, applied inside a `text(...) { }` block. */
@PoiDsl
public class RunBuilder internal constructor() {
    public var bold: Boolean? = null
    public var italic: Boolean? = null
    public var underline: Boolean? = null
    public var strikethrough: Boolean? = null

    /** Font size in points. */
    public var size: Number? = null

    /** Text color as hex, e.g. `"#C00000"`. */
    public var color: String? = null

    /** Font name, e.g. `"Calibri"`. */
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
