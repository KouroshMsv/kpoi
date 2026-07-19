package io.github.kouroshmsv.kpoi.common

import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertThrows
import org.junit.jupiter.api.Test

class RgbTest {

    @Test
    fun `parses hex with and without hash`() {
        assertEquals(Rgb(68, 114, 196), Rgb.parse("#4472C4"))
        assertEquals(Rgb(68, 114, 196), Rgb.parse("4472c4"))
    }

    @Test
    fun `formats back to hex`() {
        assertEquals("4472C4", Rgb.parse("#4472C4").toHexString())
    }

    @Test
    fun `converts to awt and bytes`() {
        val rgb = Rgb.parse("#FF8000")
        assertEquals(java.awt.Color(255, 128, 0), rgb.toAwtColor())
        assertEquals(listOf(255.toByte(), 128.toByte(), 0.toByte()), rgb.toByteArray().toList())
    }

    @Test
    fun `rejects malformed input`() {
        assertThrows(IllegalArgumentException::class.java) { Rgb.parse("#12345") }
        assertThrows(IllegalArgumentException::class.java) { Rgb.parse("blue") }
        assertThrows(IllegalArgumentException::class.java) { Rgb.parse("#GGGGGG") }
    }
}
