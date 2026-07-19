package io.github.kouroshmsv.kpoi.slides

import io.github.kouroshmsv.kpoi.common.PoiDsl
import io.github.kouroshmsv.kpoi.common.Rgb
import org.apache.poi.sl.usermodel.TextParagraph
import org.apache.poi.xslf.usermodel.XSLFTextParagraph
import org.apache.poi.xslf.usermodel.XSLFTextRun
import org.apache.poi.xslf.usermodel.XSLFTextShape

/** DSL scope for a text shape (text box or placeholder). */
@PoiDsl
public class TextShapeBuilder internal constructor(
    /** The underlying POI shape, for anything the DSL does not cover. */
    public val poiShape: XSLFTextShape,
) {
    /**
     * Adds a paragraph to the shape. When [text] is given it becomes the
     * paragraph's first run; further runs and formatting are added in [block].
     *
     * @param text optional text for the paragraph's first run.
     * @param block configures the paragraph (alignment, bullet, more runs)
     *   through the [TextParagraphBuilder] receiver.
     * @return the created [XSLFTextParagraph].
     */
    public fun paragraph(text: String? = null, block: TextParagraphBuilder.() -> Unit = {}): XSLFTextParagraph {
        val poiParagraph = poiShape.addNewTextParagraph()
        val builder = TextParagraphBuilder(poiParagraph)
        if (text != null) {
            builder.text(text)
        }
        builder.block()
        builder.applyProperties()
        return poiParagraph
    }

    /**
     * Adds one bulleted paragraph per item — a shortcut for repeated
     * `paragraph(item) { bullet = true }` calls.
     *
     * @param items the bullet lines, in order.
     */
    public fun bullets(vararg items: String) {
        items.forEach { item ->
            paragraph(item) { bullet = true }
        }
    }
}

/** DSL scope for one paragraph inside a text shape. */
@PoiDsl
public class TextParagraphBuilder internal constructor(
    /** The underlying POI paragraph, for anything the DSL does not cover. */
    public val poiParagraph: XSLFTextParagraph,
) {
    /** Whether to show a bullet in front of this paragraph. `null` leaves the layout default. */
    public var bullet: Boolean? = null

    /** Indent level `0`..`4` for nested bullets. `null` leaves the layout default. */
    public var indentLevel: Int? = null

    /**
     * Horizontal alignment of the paragraph, e.g. `TextParagraph.TextAlign.CENTER`.
     * `null` leaves the layout default (usually left-aligned).
     */
    public var align: TextParagraph.TextAlign? = null

    /**
     * Adds a run of [value] to the paragraph, formatted by [block]. Call it
     * more than once to mix formatting within a single paragraph.
     *
     * @param value the run's text.
     * @param block character formatting through the [TextRunBuilder] receiver.
     * @return the created [XSLFTextRun].
     */
    public fun text(value: String, block: TextRunBuilder.() -> Unit = {}): XSLFTextRun {
        val run = poiParagraph.addNewTextRun()
        run.setText(value)
        TextRunBuilder().apply(block).applyTo(run)
        return run
    }

    internal fun applyProperties() {
        bullet?.let { poiParagraph.isBullet = it }
        indentLevel?.let { poiParagraph.indentLevel = it }
        align?.let { poiParagraph.textAlign = it }
    }
}

/**
 * Character formatting for a single run, applied inside a `text(...) { }` block
 * (and by [SlideBuilder.title]). Every property is `null` by default, which
 * leaves the corresponding attribute at the template/layout default.
 */
@PoiDsl
public class TextRunBuilder internal constructor() {
    /** Renders the run in bold. */
    public var bold: Boolean? = null

    /** Renders the run in italic. */
    public var italic: Boolean? = null

    /** Underlines the run. */
    public var underline: Boolean? = null

    /** Font size in points, e.g. `24`. */
    public var size: Number? = null

    /** Text color as a hex string — `"#RRGGBB"` or `"RRGGBB"`, case-insensitive (e.g. `"#FFFFFF"`). */
    public var color: String? = null

    /** Font name, e.g. `"Calibri"`. */
    public var fontFamily: String? = null

    internal fun applyTo(run: XSLFTextRun) {
        bold?.let { run.isBold = it }
        italic?.let { run.isItalic = it }
        underline?.let { run.setUnderlined(it) }
        size?.let { run.fontSize = it.toDouble() }
        color?.let { run.setFontColor(Rgb.parse(it).toAwtColor()) }
        fontFamily?.let { run.fontFamily = it }
    }
}
