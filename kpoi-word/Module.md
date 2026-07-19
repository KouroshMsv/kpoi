# Module kpoi-word

`kpoi-word` is a small, type-safe Kotlin DSL over the Apache POI `XWPF` API for
building Word `.docx` documents. Instead of wiring up paragraphs, runs, tables,
and rows by hand, you describe the document with nested builder lambdas and let
the DSL create the POI objects for you.

The mental model mirrors the Word document tree:

```
document
├─ paragraph
│  └─ run          a styled span of text
└─ table
   └─ row
      └─ cell       filled like a mini-paragraph
```

- A **document** holds paragraphs and tables — `document { }` (in memory) or
  `docx(path) { }` (straight to a file).
- A **paragraph** holds one or more **runs**; each run (`text("...") { }`)
  carries its own bold / italic / underline / strikethrough / size / color /
  font formatting.
- A **table** holds **rows**, each row holds **cells**, and a cell is filled
  just like a paragraph — so the same run formatting is available inside it.

`heading(level)` is currently a convenience over a bold, sized run rather than a
real Word heading style: it emits a single bold run sized by level — **20 pt**
for level 1, **16 pt** for level 2, **14 pt** for level 3, and **12 pt** for
level 4 (levels outside `1..4` are clamped into that range). Native heading
styles are on the roadmap.

Every scope builder exposes the raw POI object it wraps —
`DocumentBuilder.poiDocument`, `ParagraphBuilder.poiParagraph`,
`TableBuilder.poiTable`, and `TableRowBuilder.poiRow` — so you can drop down to
the full POI API for anything the DSL does not cover. (`RunBuilder` is the
exception: it applies formatting onto the run, and `text(...)` returns the
underlying `XWPFRun` if you need a handle to it.)

# Package io.github.kouroshmsv.kpoi.word

The word-processing DSL: top-level document builders plus the `DocumentBuilder`,
`ParagraphBuilder`, `RunBuilder`, `TableBuilder`, and `TableRowBuilder` scopes.

## Creating and writing a document

There are two entry points:

- `document { }` builds an `XWPFDocument` in memory and hands it back to you.
  The document is `AutoCloseable`; you own it, so close it (typically with
  `use { }`) once you have written it out.
- `docx(path) { }` is the one-liner: it builds the document, writes it to
  `path`, and closes it for you.

Once you hold an `XWPFDocument` you can serialize it with the extension helpers:
`writeTo(OutputStream)` (the stream is left open), `writeTo(Path)`,
`writeTo(File)`, or `toByteArray()` for the raw `.docx` bytes. None of the
`writeTo` overloads nor `toByteArray()` close the document.

```kotlin
import io.github.kouroshmsv.kpoi.word.document
import io.github.kouroshmsv.kpoi.word.writeTo
import java.nio.file.Path

// Build in memory, then write and close explicitly.
document {
    paragraph("Hello, world")
}.use { doc ->
    doc.writeTo(Path.of("hello.docx"))
}
```

```kotlin
import io.github.kouroshmsv.kpoi.word.docx
import java.nio.file.Path

// One-liner: builds, writes to the path, and closes for you.
docx(Path.of("hello.docx")) {
    paragraph("Hello, world")
}
```

```kotlin
import io.github.kouroshmsv.kpoi.word.document
import io.github.kouroshmsv.kpoi.word.toByteArray

// In-memory bytes, e.g. for an HTTP response body.
val bytes: ByteArray = document {
    paragraph("Hello, world")
}.use { it.toByteArray() }
```

## Paragraphs and runs

`paragraph(text?) { }` adds a paragraph. The optional `text` argument becomes the
paragraph's first run; inside the block you add more runs and set paragraph-level
properties. A `RunBuilder` inside `text("...") { }` styles a single run.

Paragraph-level properties on `ParagraphBuilder`:

- `alignment: ParagraphAlignment?` — e.g. `ParagraphAlignment.CENTER` or
  `ParagraphAlignment.BOTH`; null leaves Word's default (left).
- `spacingBeforePoints: Int?` / `spacingAfterPoints: Int?` — spacing in points;
  null leaves POI's default.

Run-level properties on `RunBuilder` (each defaults to `null`, meaning "leave
POI's default untouched"):

- `bold`, `italic`, `underline`, `strikethrough`: `Boolean?`
- `size: Number?` — font size in points.
- `color: String?` — an `RRGGBB` hex string, with or without a leading `#` and
  case-insensitive, e.g. `"#C00000"` or `"c00000"`.
- `fontFamily: String?` — e.g. `"Calibri"`.

A run built with `text(value)` splits `value` on `\n`: each newline becomes a
line break (`<w:br/>`) inside that same run, so the whole string remains one
styled run. `lineBreak()` adds a standalone break as its own empty run.

```kotlin
// One paragraph, two runs — the second is bold and red.
paragraph {
    text("Hello, ")
    text("world") {
        bold = true
        color = "#FF0000"
    }
}
```

```kotlin
// Paragraph-level alignment and spacing.
paragraph("Results were strong across all three regions.") {
    alignment = ParagraphAlignment.BOTH
    spacingAfterPoints = 12
}
```

```kotlin
// Embedded newlines become line breaks within a single run.
paragraph {
    text("first line\nsecond line")
}
```

```kotlin
// The full run-formatting surface.
paragraph {
    text("styled") {
        bold = true
        italic = true
        underline = true
        strikethrough = true
        size = 14
        color = "#C00000"      // RRGGBB; the leading # is optional
        fontFamily = "Calibri"
    }
}
```

Headings are a shortcut for a bold, sized run:

```kotlin
heading("Quarterly Letter")          // level 1: bold, 20 pt
heading("Section", level = 2)        // bold, 16 pt
heading("Note", level = 3) {         // bold, 14 pt, plus extra formatting
    color = "#C00000"
}
```

## Tables

`table(width?) { }` adds a table. `width` is optional and accepts a percentage
like `"100%"`, the literal `"auto"`, or a twips value like `"5000"`; when omitted
POI's default width is used.

Inside the table, `row { }` adds a row and `cell(text?) { }` adds a cell. Both
follow a fill-then-append rule: a freshly created POI table already has one row
containing one cell, so the first `row` and first `cell` fill those pre-existing
slots before further calls create new ones. `cell(...)` returns the
`XWPFTableCell`, and its block runs against the cell's first paragraph as a
`ParagraphBuilder`, so you can align it or add styled runs.

```kotlin
table(width = "100%") {
    row {
        cell("Metric")
        cell("Value")
    }
    row {
        cell("Revenue")
        cell("$12M")
    }
}
```

```kotlin
// A cell is a paragraph: format its first paragraph via the block.
table(width = "auto") {
    row {
        cell {
            text("Header") { bold = true }
        }
        cell("Plain")
    }
}
```

## Images

`picture(path, widthPoints, heightPoints)` embeds the image at `path` as a new
run, scaled to the given width and height in points (converted to EMUs for POI).
The picture type is inferred from the file extension. Supported extensions are
`png`, `jpg`, `jpeg`, `gif`, and `bmp`, matched case-insensitively; any other
extension throws `IllegalArgumentException`.

```kotlin
import java.nio.file.Path

paragraph {
    picture(Path.of("logo.png"), widthPoints = 120, heightPoints = 40)
}
```

## End-to-end example

```kotlin
import io.github.kouroshmsv.kpoi.word.docx
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import java.nio.file.Path

fun writeLetter(target: Path) {
    docx(target) {
        heading("Quarterly Letter", level = 1)
        paragraph {
            text("Dear ")
            text("shareholders") { bold = true }
            text(",")
        }
        paragraph("Results were strong across all three regions.") {
            alignment = ParagraphAlignment.BOTH
            spacingAfterPoints = 12
        }
        table(width = "100%") {
            row {
                cell("Metric")
                cell("Value")
            }
            row {
                cell("Revenue")
                cell("$12M")
            }
            row {
                cell("NPS")
                cell("61")
            }
        }
        pageBreak()
        paragraph("Appendix follows.")
    }
}
```
