package dev.kpoi.spreadsheet

import dev.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment

/** Immutable description of a font, used as a cache key for POI fonts. */
public data class FontSpec(
    public val bold: Boolean? = null,
    public val italic: Boolean? = null,
    public val underline: Boolean? = null,
    public val strikeout: Boolean? = null,
    public val sizePoints: Double? = null,
    public val name: String? = null,
    public val color: IndexedColors? = null,
    public val colorHex: String? = null,
) {
    internal fun mergedWith(override: FontSpec): FontSpec {
        val overridesColor = override.color != null || override.colorHex != null
        return FontSpec(
            bold = override.bold ?: bold,
            italic = override.italic ?: italic,
            underline = override.underline ?: underline,
            strikeout = override.strikeout ?: strikeout,
            sizePoints = override.sizePoints ?: sizePoints,
            name = override.name ?: name,
            color = if (overridesColor) override.color else color,
            colorHex = if (overridesColor) override.colorHex else colorHex,
        )
    }
}

/** One border edge: line style plus optional indexed color. */
public data class BorderSpec(
    public val style: BorderStyle,
    public val color: IndexedColors? = null,
)

/**
 * Immutable description of a cell style, used as a cache key so that all
 * cells sharing a look also share a single POI
 * [org.apache.poi.ss.usermodel.CellStyle].
 */
public data class StyleSpec(
    public val font: FontSpec? = null,
    public val fillHex: String? = null,
    public val fillIndexed: IndexedColors? = null,
    public val dataFormat: String? = null,
    public val horizontal: HorizontalAlignment? = null,
    public val vertical: VerticalAlignment? = null,
    public val wrapText: Boolean? = null,
    public val borderTop: BorderSpec? = null,
    public val borderBottom: BorderSpec? = null,
    public val borderLeft: BorderSpec? = null,
    public val borderRight: BorderSpec? = null,
) {
    /** Returns a copy where every property set in [override] wins over this spec. */
    public fun mergedWith(override: StyleSpec): StyleSpec {
        val overridesFill = override.fillHex != null || override.fillIndexed != null
        return StyleSpec(
            font = when {
                override.font == null -> font
                font == null -> override.font
                else -> font.mergedWith(override.font)
            },
            fillHex = if (overridesFill) override.fillHex else fillHex,
            fillIndexed = if (overridesFill) override.fillIndexed else fillIndexed,
            dataFormat = override.dataFormat ?: dataFormat,
            horizontal = override.horizontal ?: horizontal,
            vertical = override.vertical ?: vertical,
            wrapText = override.wrapText ?: wrapText,
            borderTop = override.borderTop ?: borderTop,
            borderBottom = override.borderBottom ?: borderBottom,
            borderLeft = override.borderLeft ?: borderLeft,
            borderRight = override.borderRight ?: borderRight,
        )
    }

    internal val isEmpty: Boolean get() = this == EMPTY

    public companion object {
        internal val EMPTY: StyleSpec = StyleSpec()
    }
}

/**
 * A reusable, cheap style reference returned by [WorkbookBuilder.style].
 * The POI style object is only created when a cell actually uses it.
 */
public class CellStyleHandle internal constructor(internal val spec: StyleSpec)

/** Builds a [StyleSpec] from DSL configuration. */
@PoiDsl
public class StyleBuilder internal constructor() {
    /** Excel number format, e.g. `"#,##0.00"` or `"yyyy-mm-dd"`. */
    public var format: String? = null

    /** Wrap long text onto multiple lines. */
    public var wrapText: Boolean? = null

    private var font: FontSpec? = null
    private var fillHex: String? = null
    private var fillIndexed: IndexedColors? = null
    private var horizontal: HorizontalAlignment? = null
    private var vertical: VerticalAlignment? = null
    private var borderTop: BorderSpec? = null
    private var borderBottom: BorderSpec? = null
    private var borderLeft: BorderSpec? = null
    private var borderRight: BorderSpec? = null

    public fun font(block: FontBuilder.() -> Unit) {
        val built = FontBuilder().apply(block).build()
        font = font?.mergedWith(built) ?: built
    }

    /** Solid background fill from a hex color like `"#4472C4"` (XLSX only). */
    public fun fill(hex: String) {
        fillHex = hex
        fillIndexed = null
    }

    /** Solid background fill from the classic indexed palette (XLS and XLSX). */
    public fun fill(color: IndexedColors) {
        fillIndexed = color
        fillHex = null
    }

    public fun align(
        horizontal: HorizontalAlignment? = null,
        vertical: VerticalAlignment? = null,
    ) {
        horizontal?.let { this.horizontal = it }
        vertical?.let { this.vertical = it }
    }

    /** Applies the same border to all four edges. */
    public fun border(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        val spec = BorderSpec(style, color)
        borderTop = spec
        borderBottom = spec
        borderLeft = spec
        borderRight = spec
    }

    public fun borderTop(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderTop = BorderSpec(style, color)
    }

    public fun borderBottom(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderBottom = BorderSpec(style, color)
    }

    public fun borderLeft(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderLeft = BorderSpec(style, color)
    }

    public fun borderRight(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderRight = BorderSpec(style, color)
    }

    internal fun build(): StyleSpec = StyleSpec(
        font = font,
        fillHex = fillHex,
        fillIndexed = fillIndexed,
        dataFormat = format,
        horizontal = horizontal,
        vertical = vertical,
        wrapText = wrapText,
        borderTop = borderTop,
        borderBottom = borderBottom,
        borderLeft = borderLeft,
        borderRight = borderRight,
    )
}

/** Builds a [FontSpec] from DSL configuration. */
@PoiDsl
public class FontBuilder internal constructor() {
    public var bold: Boolean? = null
    public var italic: Boolean? = null
    public var underline: Boolean? = null
    public var strikeout: Boolean? = null

    /** Font size in points; fractions like 10.5 are supported. */
    public var size: Number? = null

    /** Font name, e.g. `"Calibri"`. */
    public var name: String? = null

    private var indexedColor: IndexedColors? = null
    private var hexColor: String? = null

    /** Font color from the classic indexed palette (XLS and XLSX). */
    public fun color(color: IndexedColors) {
        indexedColor = color
        hexColor = null
    }

    /** Font color from a hex string like `"#FF0000"` (XLSX only). */
    public fun color(hex: String) {
        hexColor = hex
        indexedColor = null
    }

    internal fun build(): FontSpec = FontSpec(
        bold = bold,
        italic = italic,
        underline = underline,
        strikeout = strikeout,
        sizePoints = size?.toDouble(),
        name = name,
        color = indexedColor,
        colorHex = hexColor,
    )
}
