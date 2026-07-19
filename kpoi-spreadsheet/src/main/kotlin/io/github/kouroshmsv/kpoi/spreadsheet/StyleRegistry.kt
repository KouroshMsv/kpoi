package io.github.kouroshmsv.kpoi.spreadsheet

import io.github.kouroshmsv.kpoi.common.Rgb
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont

/**
 * Materializes [StyleSpec]s into POI [CellStyle]s, creating each distinct
 * style and font at most once per workbook.
 *
 * Excel limits a workbook to 64 000 cell styles (4 000 for XLS), so the naive
 * one-style-per-cell approach breaks on real-world data sets. Keying styles
 * by their immutable spec keeps the style table minimal no matter how many
 * cells share a look.
 */
internal class StyleRegistry(private val workbook: Workbook) {
    private val styles = HashMap<StyleSpec, CellStyle>()
    private val fonts = HashMap<FontSpec, Font>()
    private val dataFormat: DataFormat by lazy { workbook.creationHelper.createDataFormat() }

    fun resolve(spec: StyleSpec): CellStyle = styles.getOrPut(spec) { materializeStyle(spec) }

    private fun materializeStyle(spec: StyleSpec): CellStyle {
        val style = workbook.createCellStyle()
        spec.font?.let { style.setFont(fonts.getOrPut(it) { materializeFont(it) }) }
        spec.dataFormat?.let { style.dataFormat = dataFormat.getFormat(it) }
        when {
            spec.fillHex != null -> {
                val xssfStyle = style as? XSSFCellStyle ?: hexColorsUnsupported()
                xssfStyle.setFillForegroundColor(XSSFColor(Rgb.parse(spec.fillHex).toByteArray(), null))
                style.fillPattern = FillPatternType.SOLID_FOREGROUND
            }
            spec.fillIndexed != null -> {
                style.fillForegroundColor = spec.fillIndexed.index
                style.fillPattern = FillPatternType.SOLID_FOREGROUND
            }
        }
        spec.horizontal?.let { style.alignment = it }
        spec.vertical?.let { style.verticalAlignment = it }
        spec.wrapText?.let { style.wrapText = it }
        spec.borderTop?.let {
            style.borderTop = it.style
            it.color?.let { color -> style.topBorderColor = color.index }
        }
        spec.borderBottom?.let {
            style.borderBottom = it.style
            it.color?.let { color -> style.bottomBorderColor = color.index }
        }
        spec.borderLeft?.let {
            style.borderLeft = it.style
            it.color?.let { color -> style.leftBorderColor = color.index }
        }
        spec.borderRight?.let {
            style.borderRight = it.style
            it.color?.let { color -> style.rightBorderColor = color.index }
        }
        return style
    }

    private fun materializeFont(spec: FontSpec): Font {
        val font = workbook.createFont()
        spec.bold?.let { font.bold = it }
        spec.italic?.let { font.italic = it }
        spec.underline?.let { font.underline = if (it) Font.U_SINGLE else Font.U_NONE }
        spec.strikeout?.let { font.strikeout = it }
        spec.sizePoints?.let { font.fontHeight = (it * TWIPS_PER_POINT).toInt().toShort() }
        spec.name?.let { font.fontName = it }
        when {
            spec.colorHex != null -> {
                val xssfFont = font as? XSSFFont ?: hexColorsUnsupported()
                xssfFont.setColor(XSSFColor(Rgb.parse(spec.colorHex).toByteArray(), null))
            }
            spec.color != null -> font.color = spec.color.index
        }
        return font
    }

    private fun hexColorsUnsupported(): Nothing = throw UnsupportedOperationException(
        "Hex colors are only supported in XLSX workbooks. For XLS, use the IndexedColors " +
            "overloads instead, e.g. fill(IndexedColors.LIGHT_BLUE) or color(IndexedColors.RED)."
    )

    private companion object {
        private const val TWIPS_PER_POINT = 20
    }
}
