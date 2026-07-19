# Module kpoi-slides

A small, type-safe Kotlin DSL for building Apache POI PowerPoint (`.pptx`)
presentations without touching the verbose XSLF API directly.

The DSL mirrors the structure of a slide deck, so builders nest the way the
content does:

```
presentation → slide → text shape (text box / placeholder) → paragraph → run
```

You open a **presentation**, add **slides** to it, drop **text shapes** (free
text boxes or layout placeholders) and **pictures** onto each slide, fill the
shapes with **paragraphs**, and format individual **runs** of text inside those
paragraphs.

A slide can be created two ways:

- **blank** — an empty canvas with no placeholders, where you position every
  shape yourself; or
- **from a `SlideLayout`** (for example `SlideLayout.TITLE_AND_CONTENT`) — so it
  inherits that layout's placeholders and the template's typography, which you
  then fill by title and index.

All coordinates and sizes — anchors, widths, heights, and font sizes — are in
**points** (1/72 inch), matching POI's own slide geometry. The default deck is
4:3; `widescreen()` switches it to 16:9 (960×540 pt).

Every builder exposes the raw POI object it wraps as an escape hatch, so you can
always reach past the DSL for anything it does not cover:
`PresentationBuilder.poiSlideShow`, `SlideBuilder.poiSlide`,
`TextShapeBuilder.poiShape`, and `TextParagraphBuilder.poiParagraph`.

# Package io.github.kouroshmsv.kpoi.slides

The entire slides DSL lives in this package: the top-level entry points
(`presentation`, `pptx`) and output helpers (`writeTo`, `toByteArray`) in
`Presentations.kt`, the `PresentationBuilder`/`SlideBuilder` scopes in
`SlideBuilder.kt`, and the text builders (`TextShapeBuilder`,
`TextParagraphBuilder`, `TextRunBuilder`) in `TextBuilders.kt`.

## Creating and writing a presentation

There are two entry points.

`pptx(path) { … }` is the one-liner: it builds the deck, writes it to the file,
and closes the underlying `XMLSlideShow` for you. Reach for it when you just want
a file on disk.

```kotlin
import io.github.kouroshmsv.kpoi.slides.pptx
import org.apache.poi.xslf.usermodel.SlideLayout
import java.nio.file.Path

pptx(Path.of("deck.pptx")) {
    slide(SlideLayout.TITLE_ONLY) {
        title("Hello")
    }
}
```

`presentation { … }` returns the open `XMLSlideShow` and hands ownership to you.
Close it (a `use { }` block is the easy way) so POI releases its resources. Use
this when you need the slide show object itself — for example to write it more
than once, or to serialize it to bytes.

```kotlin
import io.github.kouroshmsv.kpoi.slides.presentation
import io.github.kouroshmsv.kpoi.slides.toByteArray
import io.github.kouroshmsv.kpoi.slides.writeTo
import java.nio.file.Path

presentation {
    widescreen()
    slide {
        textBox(x = 0.0, y = 0.0, width = 300.0, height = 80.0) {
            paragraph("Plain slide")
        }
    }
}.use { show ->
    show.writeTo(Path.of("deck.pptx"))          // to a file
    val bytes: ByteArray = show.toByteArray()   // or to memory, e.g. an HTTP response
}
```

**Page size.** Every deck starts 4:3. Call `widescreen()` for 16:9 (960×540 pt),
or `pageSize(width, height)` for any custom size in points.

```kotlin
presentation {
    widescreen()                                 // 16:9, 960×540 pt
    // or set an explicit size in points:
    pageSize(widthPoints = 720, heightPoints = 540)   // 4:3
}
```

**Output helpers.** `toByteArray()` renders the deck to a `.pptx` byte array.
`writeTo` has three overloads:

- `writeTo(path: Path)` and `writeTo(file: File)` create (or truncate) the target
  and manage the output stream for you.
- `writeTo(out: OutputStream)` writes to a stream you own and leaves it **open**,
  so you remain responsible for closing it.

## Slides and layouts

`slide(layout) { … }` adds a slide. Without a layout it is blank — no
placeholders, so you place every shape by hand with `textBox` and `picture`.
With a `SlideLayout` it inherits that layout's placeholders and typography, which
you fill by `title(...)` and `placeholder(index) { ... }`.

```kotlin
presentation {
    // Blank slide: position everything yourself.
    slide {
        textBox(x = 36.0, y = 36.0, width = 888.0, height = 100.0) {
            paragraph("Free-form content")
        }
    }

    // From a layout: inherits the title and body placeholders.
    slide(SlideLayout.TITLE_AND_CONTENT) {
        title("Agenda")           // fills the layout's title placeholder
        placeholder(1) {          // the body/content placeholder
            bullets("Intro", "Demo", "Q&A")
        }
    }
}
```

Requesting a layout the template does not define throws
`IllegalArgumentException`, and the message lists the layouts that are available.

**Titles.** `title(text)` prefers the layout's title placeholder (POI
`Placeholder.TITLE` or `Placeholder.CENTERED_TITLE`), keeping the template's
styling. If the slide has no title placeholder — a blank slide, for instance — it
falls back to creating a bold 32-pt text box across the top of the slide. An
optional trailing block applies run formatting and can override the fallback's
bold/size:

```kotlin
slide(SlideLayout.TITLE_ONLY) {
    title("Q3 Results") {
        color = "#1F4E79"
    }
}
```

**Placeholders.** `placeholder(index)` fills the index-th placeholder inherited
from the layout (indexing is zero-based; commonly `0` is the title and `1` the
body). The slide must have been created from a layout that defines that
placeholder — calling `placeholder` on a slide that lacks it (for example a blank
slide) throws `IllegalArgumentException`.

## Text

Text lives inside a *text shape* — either a free-form `textBox` or an inherited
`placeholder`. Both open a `TextShapeBuilder`, where you add paragraphs and
bullets.

`textBox(x, y, width, height) { … }` creates a box at an explicit position and
size, in points. Inside it, `paragraph(...)` adds a paragraph and `bullets(...)`
adds one bulleted paragraph per argument:

```kotlin
textBox(x = 50.0, y = 120.0, width = 500.0, height = 200.0) {
    paragraph {
        text("Revenue up 20%") {
            bold = true
            size = 24
        }
    }
    bullets("Faster onboarding", "New regions")
}
```

`paragraph(text)` uses its argument as the paragraph's first run; the trailing
block lets you add more runs and set paragraph-level properties. Call `text(...)`
several times to mix formatting within one paragraph. Paragraph properties are
`bullet`, `indentLevel` (0..4, for nested bullets), and `align` — of type
`org.apache.poi.sl.usermodel.TextParagraph.TextAlign`. Each is `null` by default,
leaving the layout default in place.

```kotlin
import org.apache.poi.sl.usermodel.TextParagraph

textBox(x = 50.0, y = 60.0, width = 500.0, height = 260.0) {
    paragraph {
        text("Mixed ") { bold = true }
        text("formatting") { italic = true; color = "#C00000" }
    }
    paragraph("Centered heading") {
        align = TextParagraph.TextAlign.CENTER
    }
    bullets("Top-level point", "Another point")
    paragraph("Nested detail") {
        bullet = true
        indentLevel = 1
    }
}
```

**Run formatting** is set through `TextRunBuilder` inside any `text(...) { }`
block (and inside `title(...) { }`). Every property is `null` by default, which
inherits the template/layout value:

| Property     | Type      | Notes                                                      |
|--------------|-----------|------------------------------------------------------------|
| `bold`       | `Boolean?`| Bold.                                                      |
| `italic`     | `Boolean?`| Italic.                                                    |
| `underline`  | `Boolean?`| Underline.                                                 |
| `size`       | `Number?` | Font size in points, e.g. `24`.                            |
| `color`      | `String?` | Hex color, `"#RRGGBB"` or `"RRGGBB"` (case-insensitive).   |
| `fontFamily` | `String?` | Font name, e.g. `"Calibri"`.                               |

```kotlin
paragraph {
    text("Highlighted") {
        bold = true
        underline = true
        size = 18
        color = "#FFFFFF"
        fontFamily = "Calibri"
    }
}
```

## Pictures

`picture(...)` adds an image to the current slide. Position and size are in
points, and each of `x`, `y`, `width`, and `height` is optional — omit any and
the picture keeps its natural placement or size for that dimension.

Supply the bytes directly along with a `PictureType`
(`org.apache.poi.sl.usermodel.PictureData.PictureType`, defaulting to `PNG`):

```kotlin
import org.apache.poi.sl.usermodel.PictureData.PictureType

slide {
    picture(logoBytes, type = PictureType.PNG, x = 40.0, y = 40.0)
}
```

Or pass a `Path` and let the type be inferred from the file extension. Supported
extensions are `png`, `jpg`/`jpeg`, `gif`, and `bmp` (case-insensitive); any
other extension throws `IllegalArgumentException`.

```kotlin
import java.nio.file.Path

slide {
    picture(Path.of("chart.png"), x = 590.0, y = 140.0, width = 300.0, height = 168.0)
}
```

## End-to-end example

A complete deck: a widescreen title slide with formatted text, bullets, and an
image, followed by a content slide that fills an inherited body placeholder.

```kotlin
import io.github.kouroshmsv.kpoi.slides.pptx
import org.apache.poi.xslf.usermodel.SlideLayout
import java.nio.file.Path

pptx(Path.of("q3-results.pptx")) {
    widescreen()                                     // 16:9, 960×540 pt

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
        picture(Path.of("chart.png"), x = 590.0, y = 140.0, width = 300.0, height = 168.0)
    }

    slide(SlideLayout.TITLE_AND_CONTENT) {
        title("Next quarter")
        placeholder(1) {
            paragraph("Ship v1.0")
            paragraph("Grow the community")
        }
    }
}
```
