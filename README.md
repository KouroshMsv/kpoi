# kpoi — a Kotlin DSL for Apache POI

[![build](https://github.com/KouroshMsv/kpoi/actions/workflows/build.yml/badge.svg)](https://github.com/KouroshMsv/kpoi/actions/workflows/build.yml)
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

## Modules

| Module             | Wraps                          | Use it for                     |
|--------------------|--------------------------------|--------------------------------|
| `kpoi-spreadsheet` | `poi` + `poi-ooxml` (SS model) | `.xlsx`, `.xls`, streaming     |
| `kpoi-word`        | `poi-ooxml` (XWPF)             | `.docx`                        |
| `kpoi-slides`      | `poi-ooxml` (XSLF)             | `.pptx`                        |
| `kpoi-common`      | –                              | shared DSL infrastructure      |

Each module only pulls in what it needs; depend on the ones you use.

## Requirements

- Java 11+ (CI builds on 17 and 21)
- Apache POI 5.5.1 (declared transitively)
- Kotlin 2.x

## Installation

```kotlin
repositories { mavenCentral() }

dependencies {
    implementation("io.github.kouroshmsv:kpoi-spreadsheet:0.1.0")
    implementation("io.github.kouroshmsv:kpoi-word:0.1.0")
    implementation("io.github.kouroshmsv:kpoi-slides:0.1.0")
}
```

The first release is on its way to Maven Central; until it propagates you can
build from source with `./gradlew publishToMavenLocal` and `mavenLocal()`.

## Spreadsheets (`kpoi-spreadsheet`)

### Writing

```kotlin
import dev.kpoi.spreadsheet.*

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
import dev.kpoi.word.*

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
import dev.kpoi.slides.*

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
- Convention-plugin build and Maven Central publishing
- Dokka API docs

## Building

```bash
./gradlew build
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Bug reports and small, focused PRs are welcome.

## License

[Apache License 2.0](LICENSE).

Apache POI and the POI logo are trademarks of [The Apache Software Foundation](https://www.apache.org/).
This project is not affiliated with or endorsed by the ASF.
