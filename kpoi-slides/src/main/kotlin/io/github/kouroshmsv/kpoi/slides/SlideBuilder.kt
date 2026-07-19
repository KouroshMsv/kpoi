package io.github.kouroshmsv.kpoi.slides

import io.github.kouroshmsv.kpoi.common.PoiDsl
import org.apache.poi.sl.usermodel.PictureData.PictureType
import org.apache.poi.sl.usermodel.Placeholder
import org.apache.poi.xslf.usermodel.SlideLayout
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xslf.usermodel.XSLFPictureShape
import org.apache.poi.xslf.usermodel.XSLFSlide
import org.apache.poi.xslf.usermodel.XSLFTextBox
import org.apache.poi.xslf.usermodel.XSLFTextShape
import java.awt.Dimension
import java.awt.geom.Rectangle2D
import java.nio.file.Files
import java.nio.file.Path

/** Top-level DSL scope for building a presentation. */
@PoiDsl
public class PresentationBuilder internal constructor(
    /** The underlying POI slide show, for anything the DSL does not cover. */
    public val poiSlideShow: XMLSlideShow,
) {
    /** Switches the deck to 16:9 (960×540 pt). The default template is 4:3. */
    public fun widescreen() {
        pageSize(WIDESCREEN_WIDTH, WIDESCREEN_HEIGHT)
    }

    /** Slide size in points. */
    public fun pageSize(widthPoints: Int, heightPoints: Int) {
        poiSlideShow.pageSize = Dimension(widthPoints, heightPoints)
    }

    /**
     * Adds a slide. With a [layout], the slide inherits that layout's
     * placeholders (e.g. [SlideLayout.TITLE_ONLY]); without one it is blank.
     */
    public fun slide(layout: SlideLayout? = null, block: SlideBuilder.() -> Unit = {}): XSLFSlide {
        val poiSlide = if (layout == null) {
            poiSlideShow.createSlide()
        } else {
            val resolved = poiSlideShow.slideMasters.firstNotNullOfOrNull { it.getLayout(layout) }
                ?: throw IllegalArgumentException(
                    "Slide layout $layout not found in this template. Available: " +
                        poiSlideShow.slideMasters
                            .flatMap { master -> master.slideLayouts.mapNotNull { it.type } }
                            .distinct(),
                )
            poiSlideShow.createSlide(resolved)
        }
        SlideBuilder(this, poiSlide).block()
        return poiSlide
    }

    private companion object {
        private const val WIDESCREEN_WIDTH = 960
        private const val WIDESCREEN_HEIGHT = 540
    }
}

/** DSL scope for one slide. */
@PoiDsl
public class SlideBuilder internal constructor(
    private val presentation: PresentationBuilder,
    /** The underlying POI slide, for anything the DSL does not cover. */
    public val poiSlide: XSLFSlide,
) {
    /**
     * Sets the slide title. Uses the layout's title placeholder when present,
     * keeping the template's typography; otherwise creates a text box across
     * the top of the slide.
     */
    public fun title(text: String, block: TextRunBuilder.() -> Unit = {}) {
        val placeholder = poiSlide.placeholders.firstOrNull {
            it.textType == Placeholder.TITLE || it.textType == Placeholder.CENTERED_TITLE
        }
        if (placeholder != null) {
            val run = placeholder.setText(text)
            TextRunBuilder().apply(block).applyTo(run)
        } else {
            val boxWidth = presentation.poiSlideShow.pageSize.width - 2.0 * SIDE_MARGIN
            textBox(SIDE_MARGIN, TOP_MARGIN, boxWidth, FALLBACK_TITLE_HEIGHT) {
                paragraph {
                    text(text) {
                        bold = true
                        size = FALLBACK_TITLE_SIZE
                        block()
                    }
                }
            }
        }
    }

    /** Adds a free-form text box; coordinates and size in points. */
    public fun textBox(
        x: Double,
        y: Double,
        width: Double,
        height: Double,
        block: TextShapeBuilder.() -> Unit = {},
    ): XSLFTextBox {
        val box = poiSlide.createTextBox()
        box.anchor = Rectangle2D.Double(x, y, width, height)
        box.clearText()
        TextShapeBuilder(box).block()
        return box
    }

    /** Fills the [index]-th placeholder inherited from the slide layout. */
    public fun placeholder(index: Int, block: TextShapeBuilder.() -> Unit): XSLFTextShape {
        val shape = poiSlide.getPlaceholder(index) ?: throw IllegalArgumentException(
            "This slide has no placeholder $index (found ${poiSlide.placeholders.size}). " +
                "Create the slide from a layout, e.g. slide(SlideLayout.TITLE_AND_CONTENT) { ... }.",
        )
        shape.clearText()
        TextShapeBuilder(shape).block()
        return shape
    }

    /** Adds an image; position and size in points, defaulting to the image's natural size. */
    public fun picture(
        data: ByteArray,
        type: PictureType = PictureType.PNG,
        x: Double? = null,
        y: Double? = null,
        width: Double? = null,
        height: Double? = null,
    ): XSLFPictureShape {
        val pictureData = presentation.poiSlideShow.addPicture(data, type)
        val shape = poiSlide.createPicture(pictureData)
        if (x != null || y != null || width != null || height != null) {
            val anchor = shape.anchor
            shape.anchor = Rectangle2D.Double(
                x ?: anchor.x,
                y ?: anchor.y,
                width ?: anchor.width,
                height ?: anchor.height,
            )
        }
        return shape
    }

    /** Adds an image from [path]; the type is inferred from the file extension. */
    public fun picture(
        path: Path,
        x: Double? = null,
        y: Double? = null,
        width: Double? = null,
        height: Double? = null,
    ): XSLFPictureShape =
        picture(Files.readAllBytes(path), pictureTypeOf(path), x, y, width, height)

    private companion object {
        private const val SIDE_MARGIN = 36.0
        private const val TOP_MARGIN = 20.0
        private const val FALLBACK_TITLE_HEIGHT = 60.0
        private const val FALLBACK_TITLE_SIZE = 32

        private fun pictureTypeOf(path: Path): PictureType =
            when (val extension = path.fileName.toString().substringAfterLast('.').lowercase()) {
                "png" -> PictureType.PNG
                "jpg", "jpeg" -> PictureType.JPEG
                "gif" -> PictureType.GIF
                "bmp" -> PictureType.BMP
                else -> throw IllegalArgumentException(
                    "Unsupported image extension \"$extension\"; use png, jpg, gif, or bmp"
                )
            }
    }
}
