package io.github.kouroshmsv.kpoi.common

import java.awt.Color

/**
 * A parsed sRGB color used by the kpoi modules to bridge hex strings to the
 * different color representations Apache POI expects.
 *
 * Each Office XML format wants a color in its own shape, and `Rgb` is the single
 * intermediate value every kpoi module converts from:
 *
 * - XSSF (`.xlsx` spreadsheets) consume a raw 3-byte array — see [toByteArray].
 * - XWPF (`.docx` word-processing documents) consume an `"RRGGBB"` hex string — see [toHexString].
 * - XSLF (`.pptx` presentations) consume a [java.awt.Color] — see [toAwtColor].
 *
 * Instances are usually built from a hex string with [parse], but the primary
 * constructor is public and takes the three channel values directly. Every
 * channel is validated on construction.
 *
 * ```kotlin
 * val accent = Rgb.parse("#4472C4")   // Rgb(red = 68, green = 114, blue = 196)
 * accent.red                          // 68
 * accent.toHexString()                // "4472C4"
 * accent.toAwtColor()                 // java.awt.Color(68, 114, 196)
 * accent.toByteArray().toList()       // [68, 114, -60]  (bytes are signed)
 * ```
 *
 * @property red   the red channel, in `0..255`.
 * @property green the green channel, in `0..255`.
 * @property blue  the blue channel, in `0..255`.
 * @throws IllegalArgumentException if any channel is outside `0..255`.
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

    /**
     * Converts this color to a [java.awt.Color], the representation XSLF
     * (`.pptx` presentations) expects.
     *
     * @return an opaque [Color] with these [red], [green] and [blue] channels.
     */
    public fun toAwtColor(): Color = Color(red, green, blue)

    /**
     * Converts this color to a 3-byte `[r, g, b]` array, the representation XSSF
     * (`.xlsx` spreadsheets) expects for `XSSFColor`.
     *
     * Kotlin's [Byte] is signed, so any channel above `127` appears as a negative
     * number (for example `255` becomes `-1`); the underlying 8-bit pattern is
     * unchanged.
     *
     * @return a new 3-element [ByteArray] in red, green, blue order.
     */
    public fun toByteArray(): ByteArray = byteArrayOf(red.toByte(), green.toByte(), blue.toByte())

    /**
     * Converts this color to an uppercase `"RRGGBB"` hex string with no leading
     * `#`, the representation XWPF (`.docx` word-processing documents) expects.
     *
     * @return a six-character upper-case hexadecimal string, e.g. `"4472C4"`.
     */
    public fun toHexString(): String = "%02X%02X%02X".format(red, green, blue)

    /** Factory entry points for [Rgb]. */
    public companion object {
        /**
         * Parses a hex color string into an [Rgb].
         *
         * Accepts both `"#RRGGBB"` and `"RRGGBB"` (a single leading `#` is
         * optional) and is case-insensitive. After any `#` is dropped, the text
         * must be exactly six hexadecimal digits (`0`–`9`, `a`–`f`).
         *
         * ```kotlin
         * Rgb.parse("#4472C4")   // Rgb(red = 68, green = 114, blue = 196)
         * Rgb.parse("4472c4")    // same value — hash optional, case-insensitive
         * ```
         *
         * @param hex a color in `#RRGGBB` or `RRGGBB` form.
         * @return the parsed [Rgb].
         * @throws IllegalArgumentException if [hex] is not six hexadecimal digits
         * (optionally prefixed with a single `#`) — for example `"#12345"`,
         * `"blue"` or `"#GGGGGG"`.
         */
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
