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
    /** Adds a paragraph; [text] becomes its first run. */
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

    /** Adds one bulleted paragraph per item. */
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
    /** Show a bullet in front of this paragraph. */
    public var bullet: Boolean? = null

    /** Indent level 0..4 for nested bullets. */
    public var indentLevel: Int? = null

    public var align: TextParagraph.TextAlign? = null

    /** Adds a run of [value]. */
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

/** Character formatting for one run, applied inside a `text(...) { }` block. */
@PoiDsl
public class TextRunBuilder internal constructor() {
    public var bold: Boolean? = null
    public var italic: Boolean? = null
    public var underline: Boolean? = null

    /** Font size in points. */
    public var size: Number? = null

    /** Text color as hex, e.g. `"#FFFFFF"`. */
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
