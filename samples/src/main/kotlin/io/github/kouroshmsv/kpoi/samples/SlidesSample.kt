package io.github.kouroshmsv.kpoi.samples

import io.github.kouroshmsv.kpoi.slides.pptx
import org.apache.poi.xslf.usermodel.SlideLayout
import java.awt.Color
import java.awt.image.BufferedImage
import java.nio.file.Path
import javax.imageio.ImageIO

fun slidesSample(outputDir: Path): Path {
    // A stand-in "chart" so the picture API is exercised end to end.
    val chart = outputDir.resolve("chart.png")
    val image = BufferedImage(320, 180, BufferedImage.TYPE_INT_RGB)
    image.createGraphics().apply {
        color = Color(0x44, 0x72, 0xC4)
        fillRect(0, 0, 320, 180)
        color = Color.WHITE
        drawString("Q3", 24, 40)
        dispose()
    }
    ImageIO.write(image, "png", chart.toFile())

    val target = outputDir.resolve("deck.pptx")
    pptx(target) {
        widescreen()
        slide(SlideLayout.TITLE_ONLY) {
            title("Q3 Results")
            textBox(x = 60.0, y = 130.0, width = 480.0, height = 260.0) {
                paragraph {
                    text("Revenue up 20%") {
                        bold = true
                        size = 28
                    }
                }
                bullets("Faster onboarding", "Two new regions", "NPS at 61")
            }
            picture(chart, x = 590.0, y = 140.0, width = 300.0, height = 168.0)
        }
        slide(SlideLayout.TITLE_AND_CONTENT) {
            title("Next quarter")
            placeholder(1) {
                paragraph("Ship v1.0")
                paragraph("Grow the community")
            }
        }
    }
    return target
}
