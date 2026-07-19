package io.github.kouroshmsv.kpoi.slides

import org.apache.poi.xslf.usermodel.SlideLayout
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFTextShape
import org.junit.jupiter.api.Assertions.assertEquals
import org.junit.jupiter.api.Assertions.assertTrue
import org.junit.jupiter.api.Test
import java.io.ByteArrayInputStream

class SlidesDslTest {

    @Test
    fun `round-trips slides, titles, and text boxes`() {
        val bytes = presentation {
            widescreen()
            slide(SlideLayout.TITLE_ONLY) {
                title("Quarterly Update")
                textBox(x = 50.0, y = 120.0, width = 500.0, height = 200.0) {
                    paragraph {
                        text("Revenue up 20%") {
                            bold = true
                            size = 24
                        }
                    }
                    bullets("Faster onboarding", "New regions")
                }
            }
            slide {
                textBox(x = 10.0, y = 10.0, width = 300.0, height = 80.0) {
                    paragraph("Plain slide")
                }
            }
        }.use { it.toByteArray() }

        XMLSlideShow(ByteArrayInputStream(bytes)).use { slideShow ->
            assertEquals(960, slideShow.pageSize.width)
            assertEquals(540, slideShow.pageSize.height)
            assertEquals(2, slideShow.slides.size)

            val firstSlideTexts = slideShow.slides[0].shapes
                .filterIsInstance<XSLFTextShape>()
                .map { it.text }
            assertTrue(firstSlideTexts.any { it.contains("Quarterly Update") })
            assertTrue(firstSlideTexts.any { it.contains("Revenue up 20%") })
            assertTrue(firstSlideTexts.any { it.contains("Faster onboarding") })

            val secondSlideTexts = slideShow.slides[1].shapes
                .filterIsInstance<XSLFTextShape>()
                .map { it.text }
            assertTrue(secondSlideTexts.any { it.contains("Plain slide") })
        }
    }

    @Test
    fun `styled runs keep their formatting`() {
        presentation {
            slide {
                textBox(x = 0.0, y = 0.0, width = 200.0, height = 50.0) {
                    paragraph {
                        text("Styled") {
                            bold = true
                            size = 30
                        }
                    }
                }
            }
        }.use { slideShow ->
            val shape = slideShow.slides[0].shapes.filterIsInstance<XSLFTextShape>().first()
            val run = shape.textParagraphs[0].textRuns[0]
            assertEquals("Styled", run.rawText)
            assertTrue(run.isBold)
            assertEquals(30.0, run.fontSize)
        }
    }
}
