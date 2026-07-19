package dev.kpoi.common

import java.awt.Color

/**
 * A parsed sRGB color used by the kpoi modules to bridge hex strings to the
 * different color representations Apache POI expects (XSSF byte arrays,
 * XWPF hex strings, XSLF AWT colors).
 */
public data class Rgb(
    public val red: Int,
    public val green: Int,
    public val blue: Int,
) {
    init {
        require(red in 0..255 && green in 0..255 && blue in 0..255) {
            "Color channels must be in 0..255, got ($red, $green, $blue)"
        }
    }

    /** As [java.awt.Color], used by XSLF. */
    public fun toAwtColor(): Color = Color(red, green, blue)

    /** As a 3-byte `[r, g, b]` array, used by XSSF. */
    public fun toByteArray(): ByteArray = byteArrayOf(red.toByte(), green.toByte(), blue.toByte())

    /** As `"RRGGBB"` without a leading `#`, used by XWPF. */
    public fun toHexString(): String = "%02X%02X%02X".format(red, green, blue)

    public companion object {
        /** Parses `"#RRGGBB"` or `"RRGGBB"` (case-insensitive). */
        public fun parse(hex: String): Rgb {
            val digits = hex.removePrefix("#")
            require(digits.length == 6 && digits.all { it.isDigit() || it.lowercaseChar() in 'a'..'f' }) {
                "Expected a color in #RRGGBB format, got \"$hex\""
            }
            return Rgb(
                digits.substring(0, 2).toInt(16),
                digits.substring(2, 4).toInt(16),
                digits.substring(4, 6).toInt(16),
            )
        }
    }
}
