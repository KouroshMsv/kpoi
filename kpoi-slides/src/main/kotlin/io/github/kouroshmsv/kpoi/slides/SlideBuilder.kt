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

/**
 * Top-level DSL scope for building a presentation: set the page size, then add
 * slides. This is the receiver of the [presentation] and [pptx] builder
 * lambdas.
 *
 * The raw [XMLSlideShow] is available as [poiSlideShow] for anything the DSL
 * does not cover.
 */
@PoiDsl
public class PresentationBuilder internal constructor(
    /** The underlying POI slide show, for anything the DSL does not cover. */
    public val poiSlideShow: XMLSlideShow,
) {
    /** Switches the deck to 16:9 (960×540 pt). The default template is 4:3. */
    public fun widescreen() {
        pageSize(WIDESCREEN_WIDTH, WIDESCREEN_HEIGHT)
    }

    /**
     * Sets the slide size, in points, for custom aspect ratios; [widescreen] is
     * the 16:9 shortcut.
     *
     * @param widthPoints slide width in points.
     * @param heightPoints slide height in points.
     */
    public fun pageSize(widthPoints: Int, heightPoints: Int) {
        poiSlideShow.pageSize = Dimension(widthPoints, heightPoints)
    }

    /**
     * Adds a slide to the deck and configures it through the [SlideBuilder]
     * receiver.
     *
     * Pass a [layout] to base the slide on one of the template's layouts so it
     * inherits that layout's placeholders (title, body, ...) and typography —
     * for example [SlideLayout.TITLE_AND_CONTENT] or [SlideLayout.TITLE_ONLY].
     * Omit it (or pass `null`) to start from a blank slide with no placeholders.
     *
     * ```kotlin
     * slide(SlideLayout.TITLE_AND_CONTENT) {
     *     title("Roadmap")
     *     placeholder(1) {
     *         bullets("Ship v1.0", "Grow the community")
     *     }
     * }
     * ```
     *
     * @param layout the layout to inherit placeholders from, or `null` for a blank slide.
     * @param block configures the new slide through the [SlideBuilder] receiver.
     * @return the created [XSLFSlide].
     * @throws IllegalArgumentException if [layout] is not defined by any slide
     *   master in the template.
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

/**
 * DSL scope for one slide: set its title, add text boxes, fill inherited
 * placeholders, and place pictures. All positions and sizes are in points.
 *
 * The raw [XSLFSlide] is available as [poiSlide] for anything the DSL does not
 * cover.
 */
@PoiDsl
public class SlideBuilder internal constructor(
    private val presentation: PresentationBuilder,
    /** The underlying POI slide, for anything the DSL does not cover. */
    public val poiSlide: XSLFSlide,
) {
    /**
     * Sets the slide title. When the slide's layout provides a title
     * placeholder (POI `Placeholder.TITLE` or `Placeholder.CENTERED_TITLE`) the
     * text goes there, keeping the template's typography; otherwise a bold
     * 32-pt text box is created across the top of the slide as a fallback.
     *
     * The optional [block] applies run-level formatting on top of that, and can
     * override the fallback's bold/size.
     *
     * ```kotlin
     * slide(SlideLayout.TITLE_ONLY) {
     *     title("Q3 Results") {
     *         color = "#1F4E79"
     *     }
     * }
     * ```
     *
     * @param text the title text.
     * @param block optional character formatting for the title run.
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

    /**
     * Adds a free-form text box at the given position and size, in points, and
     * fills it through the [TextShapeBuilder] receiver (paragraphs, bullets,
     * runs). Unlike [placeholder], a text box works on any slide, blank or not.
     *
     * ```kotlin
     * textBox(x = 50.0, y = 120.0, width = 500.0, height = 200.0) {
     *     paragraph {
     *         text("Revenue up 20%") {
     *             bold = true
     *             size = 24
     *         }
     *     }
     *     bullets("Faster onboarding", "New regions")
     * }
     * ```
     *
     * @param x left edge in points.
     * @param y top edge in points.
     * @param width box width in points.
     * @param height box height in points.
     * @param block fills the box through the [TextShapeBuilder] receiver.
     * @return the created [XSLFTextBox].
     */
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

    /**
     * Fills the [index]-th placeholder inherited from the slide's layout.
     * Existing placeholder text is cleared first, then the placeholder is filled
     * through the [TextShapeBuilder] receiver.
     *
     * The slide must have been created from a layout that defines this
     * placeholder — see [PresentationBuilder.slide]. Placeholder indexing is
     * zero-based; commonly `0` is the title and `1` the body/content.
     *
     * ```kotlin
     * slide(SlideLayout.TITLE_AND_CONTENT) {
     *     placeholder(1) {
     *         bullets("First point", "Second point")
     *     }
     * }
     * ```
     *
     * @param index zero-based placeholder index within the slide's layout.
     * @param block fills the placeholder through the [TextShapeBuilder] receiver.
     * @return the filled [XSLFTextShape].
     * @throws IllegalArgumentException if the slide has no placeholder at
     *   [index] (for example, it was created blank rather than from a layout).
     */
    public fun placeholder(index: Int, block: TextShapeBuilder.() -> Unit): XSLFTextShape {
        val shape = poiSlide.getPlaceholder(index) ?: throw IllegalArgumentException(
            "This slide has no placeholder $index (found ${poiSlide.placeholders.size}). " +
                "Create the slide from a layout, e.g. slide(SlideLayout.TITLE_AND_CONTENT) { ... }.",
        )
        shape.clearText()
        TextShapeBuilder(shape).block()
        return shape
    }

    /**
     * Places an image on the slide from raw [data] of the given [type].
     * Position and size are in points; each of [x], [y], [width], and [height]
     * keeps the picture's natural placement/size when left `null`, and only the
     * values you supply override it.
     *
     * @param data the encoded image bytes.
     * @param type the image format; defaults to [PictureType.PNG].
     * @param x left edge in points, or `null` to keep the natural value.
     * @param y top edge in points, or `null` to keep the natural value.
     * @param width width in points, or `null` to keep the natural value.
     * @param height height in points, or `null` to keep the natural value.
     * @return the created [XSLFPictureShape].
     */
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

    /**
     * Places an image on the slide, reading the bytes from [path] and inferring
     * the [PictureType] from the file extension. Supported extensions are
     * `png`, `jpg`/`jpeg`, `gif`, and `bmp` (case-insensitive). Position and
     * size are in points; a `null` coordinate keeps the picture's natural value.
     *
     * ```kotlin
     * picture(Path.of("chart.png"), x = 590.0, y = 140.0, width = 300.0, height = 168.0)
     * ```
     *
     * @param path the image file to read.
     * @param x left edge in points, or `null` to keep the natural value.
     * @param y top edge in points, or `null` to keep the natural value.
     * @param width width in points, or `null` to keep the natural value.
     * @param height height in points, or `null` to keep the natural value.
     * @return the created [XSLFPictureShape].
     * @throws IllegalArgumentException if the file extension is not one of
     *   `png`, `jpg`, `jpeg`, `gif`, or `bmp`.
     */
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
                    "Unsupported image extension \"$extension\"; use png, jpg, jpeg, gif, or bmp"
                )
            }
    }
}
