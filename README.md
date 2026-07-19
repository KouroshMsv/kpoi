# kpoi — a Kotlin DSL for Apache POI

[![build](https://github.com/KouroshMsv/kpoi/actions/workflows/build.yml/badge.svg)](https://github.com/KouroshMsv/kpoi/actions/workflows/build.yml)
[![Maven Central](https://img.shields.io/maven-central/v/io.github.kouroshmsv/kpoi-spreadsheet)](https://central.sonatype.com/search?namespace=io.github.kouroshmsv)
[![Apache POI](https://img.shields.io/badge/Apache_POI-5.5.1-blue)](https://poi.apache.org/)
[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](LICENSE)

Type-safe, concise Kotlin builders for the three big [Apache POI](https://poi.apache.org/) document
families: **Excel** (XSSF / SXSSF / HSSF), **Word** (XWPF), and **PowerPoint** (XSLF).

```kotlin
xlsx(Path.of("report.xlsx")) {
    val header = style {
        font { bold = true; color("#FFFFFF") }
        fill("#4472C4")
    }
    sheet("Report") {
        freeze(rows = 1)
        row(header) { cell("Name"); cell("Score"); cell("Joined") }
        people.forEach { person ->
            row {
                cell(person.name)
                cell(person.score) { format = "0.00" }
                cell(person.joined)              // LocalDate, formatted automatically
            }
        }
        row { cell("Total"); formula("SUM(B2:B${people.size + 1})") }
        autoSizeColumns(0, 1, 2)
    }
}
```

kpoi is an independent community project, not a product of the Apache Software Foundation.
It builds **on top of** POI's public API — every builder exposes the underlying POI object
(`poiWorkbook`, `poiSheet`, `poiRun`, …), so you can always drop down to the full API.

kpoi is young (0.x): the DSL surface may still grow and shift until 1.0, so pin your
version. Since raw POI is always exposed, nothing the DSL doesn't cover can block you.

API reference: **[kouroshmsv.github.io/kpoi](https://kouroshmsv.github.io/kpoi/)**

## Modules

| Module             | Wraps                          | Use it for                     |
|--------------------|--------------------------------|--------------------------------|
| `kpoi-spreadsheet` | `poi` + `poi-ooxml` (SS model) | `.xlsx`, `.xls`, streaming     |
| `kpoi-word`        | `poi-ooxml` (XWPF)             | `.docx`                        |
| `kpoi-slides`      | `poi-ooxml` (XSLF)             | `.pptx`                        |
| `kpoi-common`      | –                              | shared DSL infrastructure      |

Each module only pulls in what it needs; depend on the ones you use.

## Compatibility

| kpoi  | Apache POI | Java baseline | Kotlin |
|-------|------------|---------------|--------|
| 0.1.x | 5.5.1      | 11+           | 2.x    |

Each kpoi release is built and tested against exactly the [Apache POI](https://poi.apache.org/)
version above (CI runs on JDK 17 and 21). POI is exposed as an `api` dependency, so forcing a
newer 5.x patch in your own build usually works — but the table is what CI actually verifies.

## Installation

```kotlin
repositories { mavenCentral() }

dependencies {
    implementation("io.github.kouroshmsv:kpoi-spreadsheet:0.1.0")
    implementation("io.github.kouroshmsv:kpoi-word:0.1.0")
    implementation("io.github.kouroshmsv:kpoi-slides:0.1.0")
}
```

## Spreadsheets (`kpoi-spreadsheet`)

### Writing

```kotlin
import io.github.kouroshmsv.kpoi.spreadsheet.*

val workbook = workbook {                    // XSSFWorkbook; also: xlsWorkbook {}, streamingWorkbook {}
    val money = style { format = "#,##0.00" }

    sheet("Q3") {
        columnWidth(0, characters = 24)
        row {
            cell("Item")
            cell("Price", money)
            cell(true)
            cell(LocalDate.now())
            hyperlink("https://poi.apache.org", "docs")
        }
        merge("A3:C3")
        autoFilter("A1:E1")
    }
}
workbook.writeTo(Path.of("q3.xlsx"))         // or .toByteArray() for HTTP responses
workbook.close()
```

- **One-liner to disk:** `xlsx(path) { ... }` handles writing and closing; pass
  `streaming = true` for constant-memory exports of millions of rows.
- **Styles deduplicate automatically.** Excel caps a workbook at 64k styles; kpoi keys every
  style by its content, so a million cells styled `{ font { bold = true } }` create exactly
  one POI `CellStyle`. Reusable handles (`val h = style { ... }`) and inline `style { ... }`
  blocks share the same cache, and row styles merge with per-cell overrides.
- **Dates just work:** `LocalDate` / `LocalDateTime` / `Date` cells get a sensible default
  format (configurable via `defaultDateFormat`) unless you set one.

### Reading helpers

```kotlin
val sheet = XSSFWorkbook(file).getSheetAt(0)
sheet["B2"]?.doubleOrNull()
sheet.cellAt("C7")?.localDateOrNull()
cell.displayString()                          // exactly as Excel renders it
```

## Word (`kpoi-word`)

```kotlin
import io.github.kouroshmsv.kpoi.word.*

docx(Path.of("letter.docx")) {
    heading("Quarterly Letter", level = 1)
    paragraph {
        text("Dear ")
        text("shareholders") { bold = true }
        text(",")
    }
    paragraph("Results were strong.") {
        alignment = ParagraphAlignment.BOTH
        spacingAfterPoints = 12
    }
    table(width = "100%") {
        row { cell("Metric"); cell("Value") }
        row { cell("Revenue"); cell("$12M") }
    }
    pageBreak()
    paragraph { picture(Path.of("chart.png"), widthPoints = 400, heightPoints = 240) }
}
```

## PowerPoint (`kpoi-slides`)

```kotlin
import io.github.kouroshmsv.kpoi.slides.*

pptx(Path.of("deck.pptx")) {
    widescreen()
    slide(SlideLayout.TITLE_ONLY) {
        title("Q3 Results")                   // uses the layout's placeholder typography
        textBox(x = 50.0, y = 120.0, width = 860.0, height = 300.0) {
            paragraph { text("Revenue up 20%") { bold = true; size = 28 } }
            bullets("Faster onboarding", "Two new regions", "NPS at 61")
        }
        picture(Path.of("chart.png"), x = 520.0, y = 140.0)
    }
}
```

## Design principles

1. **Escape hatch everywhere.** The DSL covers the common 90%; the raw POI object is one
   property away for the rest. kpoi never hides or replaces POI's model.
2. **Content-keyed style caching** instead of "create a style per cell" (the #1 POI pitfall).
3. **`@DslMarker` scoping** so builders can't leak into each other.
4. **Explicit API mode** — the public surface is deliberate, documented, and binary-friendly.
5. **Thin by design.** No reflection, no annotation processing, no runtime magic; each
   builder call maps 1:1 onto POI calls you could have written yourself.

## Roadmap

- Word heading/style templates and numbered lists
- Cell comments, images, and conditional formatting in spreadsheets
- Slide tables and notes
- Convention-plugin build

## Samples

Runnable end-to-end examples live in [`samples/`](samples):

```bash
./gradlew :samples:run   # writes report.xlsx, letter.docx, deck.pptx to samples/build/sample-output
```

## Building

```bash
./gradlew build
```

## Support

- kpoi bugs and feature requests: [GitHub issues](https://github.com/KouroshMsv/kpoi/issues) —
  please include the kpoi/POI versions and a minimal snippet.
- Questions about POI itself (file formats, behavior of the underlying API): the
  [POI components guide](https://poi.apache.org/components/) and
  [POI mailing lists](https://poi.apache.org/mailinglists.html) are the best places.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Bug reports and small, focused PRs are welcome.

## License

[Apache License 2.0](LICENSE).

Apache POI and the POI logo are trademarks of [The Apache Software Foundation](https://www.apache.org/).
This project is not affiliated with or endorsed by the ASF.

Built on [Apache POI](https://poi.apache.org/); developed with help from
[Claude Code](https://claude.com/claude-code).
