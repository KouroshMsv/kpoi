package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.VerticalAlignment

/**
 * Immutable description of a font, used as the cache key that lets identical
 * fonts share a single POI `Font`. You normally build one through [FontBuilder]
 * (via a `font { }` block) rather than constructing it directly; each `null`
 * property means "leave unset / inherit".
 *
 * @property bold render text bold.
 * @property italic render text italic.
 * @property underline single underline.
 * @property strikeout strike a line through the text.
 * @property sizePoints font size in points.
 * @property name font family name, e.g. `"Calibri"`.
 * @property color font color from the indexed palette (XLS and XLSX).
 * @property colorHex font color as `"#RRGGBB"` (XLSX only).
 */
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

/**
 * One border edge: a line [style] plus an optional indexed [color].
 *
 * @property style the border line style.
 * @property color optional edge color from the indexed palette.
 */
public data class BorderSpec(
    public val style: BorderStyle,
    public val color: IndexedColors? = null,
)

/**
 * Immutable description of a cell style, used as the cache key so that all cells
 * sharing a look also share a single POI
 * [org.apache.poi.ss.usermodel.CellStyle]. You normally build one through
 * [StyleBuilder] (via [WorkbookBuilder.style] or an inline `style { }` block);
 * each `null` property means "leave unset / inherit".
 *
 * @property font the font, or `null` for the default.
 * @property fillHex solid background fill as `"#RRGGBB"` (XLSX only).
 * @property fillIndexed solid background fill from the indexed palette (XLS and XLSX).
 * @property dataFormat Excel number format, e.g. `"#,##0.00"`.
 * @property horizontal horizontal alignment.
 * @property vertical vertical alignment.
 * @property wrapText wrap long text onto multiple lines.
 * @property borderTop top edge border.
 * @property borderBottom bottom edge border.
 * @property borderLeft left edge border.
 * @property borderRight right edge border.
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
    /**
     * Returns a copy where every property set (non-`null`) in [override] wins
     * over this spec; unset properties in [override] keep this spec's values.
     * Fill and font color are replaced as a unit, so setting a hex fill clears an
     * inherited indexed fill and vice versa.
     *
     * @param override the higher-priority spec.
     * @return the merged spec.
     */
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

/**
 * Builds a [StyleSpec] from DSL configuration; the receiver of
 * [WorkbookBuilder.style] and of inline `cell(...) { style { } }` blocks. Every
 * option left unset is inherited when styles merge.
 */
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

    /**
     * Configures the cell font. Calling this more than once merges the settings,
     * later calls overriding earlier ones.
     *
     * @param block font configuration.
     */
    public fun font(block: FontBuilder.() -> Unit) {
        val built = FontBuilder().apply(block).build()
        font = font?.mergedWith(built) ?: built
    }

    /**
     * Solid background fill from a hex color like `"#4472C4"`. Supported only in
     * XLSX; in a legacy XLS ([xlsWorkbook]/[xls]) workbook, writing a cell that
     * uses this style throws [UnsupportedOperationException] — use the [fill]
     * overload that takes an [IndexedColors] instead.
     *
     * @param hex color as `"#RRGGBB"` or `"RRGGBB"` (case-insensitive).
     */
    public fun fill(hex: String) {
        fillHex = hex
        fillIndexed = null
    }

    /**
     * Solid background fill from the classic indexed palette, e.g.
     * [IndexedColors.LIGHT_BLUE]. Works in both XLS and XLSX.
     *
     * @param color palette color.
     */
    public fun fill(color: IndexedColors) {
        fillIndexed = color
        fillHex = null
    }

    /**
     * Sets horizontal and/or vertical text alignment. Arguments left `null` are
     * unchanged.
     *
     * @param horizontal horizontal alignment, or `null` to leave unset.
     * @param vertical vertical alignment, or `null` to leave unset.
     */
    public fun align(
        horizontal: HorizontalAlignment? = null,
        vertical: VerticalAlignment? = null,
    ) {
        horizontal?.let { this.horizontal = it }
        vertical?.let { this.vertical = it }
    }

    /**
     * Applies the same border to all four edges.
     *
     * @param style border line style; defaults to [BorderStyle.THIN].
     * @param color optional border color from the indexed palette.
     */
    public fun border(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        val spec = BorderSpec(style, color)
        borderTop = spec
        borderBottom = spec
        borderLeft = spec
        borderRight = spec
    }

    /**
     * Sets the top border edge.
     *
     * @param style border line style; defaults to [BorderStyle.THIN].
     * @param color optional border color from the indexed palette.
     */
    public fun borderTop(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderTop = BorderSpec(style, color)
    }

    /**
     * Sets the bottom border edge.
     *
     * @param style border line style; defaults to [BorderStyle.THIN].
     * @param color optional border color from the indexed palette.
     */
    public fun borderBottom(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderBottom = BorderSpec(style, color)
    }

    /**
     * Sets the left border edge.
     *
     * @param style border line style; defaults to [BorderStyle.THIN].
     * @param color optional border color from the indexed palette.
     */
    public fun borderLeft(style: BorderStyle = BorderStyle.THIN, color: IndexedColors? = null) {
        borderLeft = BorderSpec(style, color)
    }

    /**
     * Sets the right border edge.
     *
     * @param style border line style; defaults to [BorderStyle.THIN].
     * @param color optional border color from the indexed palette.
     */
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

/**
 * Builds a [FontSpec] from DSL configuration; the receiver of [StyleBuilder.font].
 */
@PoiDsl
public class FontBuilder internal constructor() {
    /** Render text bold. */
    public var bold: Boolean? = null

    /** Render text italic. */
    public var italic: Boolean? = null

    /** Underline the text (single underline). */
    public var underline: Boolean? = null

    /** Strike a line through the text. */
    public var strikeout: Boolean? = null

    /** Font size in points; fractions like 10.5 are supported. */
    public var size: Number? = null

    /** Font name, e.g. `"Calibri"`. */
    public var name: String? = null

    private var indexedColor: IndexedColors? = null
    private var hexColor: String? = null

    /**
     * Font color from the classic indexed palette, e.g. [IndexedColors.RED].
     * Works in both XLS and XLSX.
     *
     * @param color palette color.
     */
    public fun color(color: IndexedColors) {
        indexedColor = color
        hexColor = null
    }

    /**
     * Font color from a hex string like `"#FF0000"`. Supported only in XLSX; in a
     * legacy XLS workbook, writing a cell that uses this style throws
     * [UnsupportedOperationException] — use the [color] overload that takes an
     * [IndexedColors] instead.
     *
     * @param hex color as `"#RRGGBB"` or `"RRGGBB"` (case-insensitive).
     */
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
