# Module kpoi-spreadsheet

A small, type-safe Kotlin DSL over [Apache POI](https://poi.apache.org/) for
producing and reading Excel spreadsheets without the ceremony of POI's imperative
API.

The module speaks three formats through one builder:

- **XLSX** (`workbook { }` / `xlsx(path) { }`) — the modern `.xlsx` format, held
  fully in memory. This is the default choice.
- **Streaming XLSX** (`streamingWorkbook { }` / `xlsx(path, streaming = true) { }`)
  — the same `.xlsx` output backed by POI's SXSSF, which flushes older rows to a
  temporary file and keeps only a small window in memory. Use it for exports large
  enough to strain the heap.
- **Legacy XLS** (`xlsWorkbook { }` / `xls(path) { }`) — the old `.xls` binary
  format, for consumers that cannot read `.xlsx`. It is more limited (65 536 rows,
  256 columns, 4 000 styles, indexed colors only).

Every format is built the same way, following a nested mental model that mirrors a
spreadsheet's structure:

```
workbook -> sheet -> row -> cell
```

You open a `workbook { }`, add one or more `sheet { }`s, append `row { }`s, and
write `cell(...)`s into each row. Styling, freezing, merging, formulas, and
hyperlinks all hang off the builder that owns them.

## The headline feature: content-keyed style deduplication

Excel caps a workbook at roughly **64 000 cell styles** (only **4 000** for legacy
XLS). The naive approach of creating one POI `CellStyle` per cell blows past that
limit on real-world data sets and corrupts the file. kpoi avoids the problem
entirely: a style is described by an immutable, value-based `StyleSpec`, and every
cell whose style resolves to the same spec **shares a single** POI `CellStyle`
(and `Font`). Declare one bold style and apply it to a million cells and the file
still contains exactly one bold style. You never think about the style table — it
stays minimal automatically.

# Package io.github.kouroshmsv.kpoi.spreadsheet

The whole public surface lives in this one package: the workbook/sheet/row/cell
builders, the styling DSL, the write helpers, and a handful of extensions for
reading values back out. Each builder also exposes an escape hatch to the
underlying POI object (`poiWorkbook`, `poiSheet`, `poiRow`, `poiCell`) for the rare
feature the DSL does not wrap.

## Creating and writing a workbook

The in-memory builders return an **open** POI workbook that you own and must close;
the file one-liners create, write, and close for you.

| Entry point | Returns / does | Format |
| --- | --- | --- |
| `workbook { }` | open `XSSFWorkbook` | XLSX |
| `streamingWorkbook(rowAccessWindowSize) { }` | open `SXSSFWorkbook` | streaming XLSX |
| `xlsWorkbook { }` | open `HSSFWorkbook` | legacy XLS |
| `xlsx(path, streaming = false) { }` | writes a file, releases resources | XLSX |
| `xls(path) { }` | writes a file, releases resources | legacy XLS |

When you hold a workbook yourself, wrap it in `use { }` and pick an output sink with
one of the `writeTo` overloads (`OutputStream`, `Path`, or `File`) or render bytes
with `toByteArray()`:

```kotlin
// In-memory, then write to a path (close with `use`).
workbook {
    sheet("Report") {
        row { cell("Hello") }
    }
}.use { wb ->
    wb.writeTo(Path.of("report.xlsx"))
}

// Or serve the bytes directly, e.g. as an HTTP response body.
val bytes: ByteArray = workbook {
    sheet { row { cell("Hello") } }
}.use { it.toByteArray() }
```

The file one-liner does the same in one call:

```kotlin
xlsx(Path.of("report.xlsx")) {
    sheet("Report") {
        row { cell("Hello") }
    }
}
```

For very large exports, switch to streaming. `xlsx(path, streaming = true)` uses an
`SXSSFWorkbook` under the hood and cleans up its temporary files on close:

```kotlin
xlsx(Path.of("huge.xlsx"), streaming = true) {
    sheet("Big") {
        repeat(1_000_000) { i ->
            row { cell("row $i"); cell(i) }
        }
    }
}
```

`writeTo(OutputStream)` flushes but does **not** close the stream — you remain
responsible for it. `writeTo(Path)` and `writeTo(File)` create or overwrite the
target file.

## Sheets

Inside a `sheet(name) { }` block you configure the sheet and add rows. The `name`
is sanitized with POI's `WorkbookUtil.createSafeSheetName` (invalid characters
replaced, trimmed to 31 characters); pass `null` to let POI assign a default name.

```kotlin
sheet("Q3 Report") {
    columnWidth(0, characters = 20)     // width in character widths
    freeze(rows = 1)                    // pin the header row while scrolling
    freeze(rows = 1, columns = 1)       // pin the first row and column
    // ... rows ...
    merge("A6:C6")                      // merge a range in A1 notation
    merge(firstRow = 5, lastRow = 5, firstColumn = 0, lastColumn = 2) // or by bounds
    autoFilter("A1:C1")                 // filter dropdowns over the header
}
```

**Auto-sizing** measures content and must be called *after* the rows are written:

```kotlin
sheet("Report") {
    row { cell("Name"); cell("Score") }
    // ... more rows ...
    autoSizeColumns(0, 1)
}
```

On a **streaming (SXSSF) sheet** the older rows are already flushed to disk and can
no longer be measured, so you must opt in *before* writing any rows, or POI throws:

```kotlin
xlsx(Path.of("big.xlsx"), streaming = true) {
    sheet("Big") {
        trackColumnsForAutoSize()       // required for SXSSF; no-op on normal sheets
        repeat(10_000) { row { cell("row $it") } }
        autoSizeColumns(0)
    }
}
```

`rowAt(index)` writes at an explicit zero-based row and resumes appends from the next
row; `skipRows(count)` leaves gaps before the next `row`.

## Rows and cells

`row(style) { }` appends a row; the optional `style` handle becomes the default for
every cell in that row. There is one `cell(...)` overload per value type, so the
right Excel cell type is chosen automatically:

```kotlin
row {
    cell("Alice")                       // String? — text
    cell(95.5)                          // Number  — stored as Double
    cell(true)                          // Boolean — Excel TRUE/FALSE
    cell(LocalDate.of(2026, 7, 19))     // date, formatted yyyy-mm-dd
    cell(LocalDateTime.now())           // date-time, formatted yyyy-mm-dd hh:mm
    cell(Date())                        // java.util.Date, as a date-time
}
```

Passing `null` to the `String?` overload creates a **styled blank cell** — useful
for keeping a fill or border in an otherwise empty position:

```kotlin
row(header) {
    cell("Name")
    cell(null)                          // blank, but still carries the header style
}
```

Two more cell kinds:

```kotlin
row {
    cell("Total")
    formula("SUM(B2:B9)")               // no leading '='; Excel computes on open
}
row {
    hyperlink("https://poi.apache.org") // blue + underlined, label defaults to the URL
    hyperlink("https://poi.apache.org", "Apache POI") // custom label
}
```

Each `cell(...)` takes a trailing configuration block exposing per-cell options:

```kotlin
row {
    cell(1234.5) { format = "#,##0.00" }            // number format for just this cell
    cell("Header") { style { font { bold = true } } } // inline one-off style
    cell("docs") { hyperlink("https://example.com") } // link without the default styling
}
```

Other row helpers: `height(points)` sets the row height, and `skipCells(count)`
leaves empty columns before the next cell.

### Default date formats

Date and date-time cells receive a default number format unless one is set on the
cell or its style. The defaults live on the workbook builder and can be changed for
the whole workbook:

```kotlin
workbook {
    defaultDateFormat = "dd/MM/yyyy"        // applied to LocalDate cells
    defaultDateTimeFormat = "yyyy-mm-dd hh:mm:ss" // applied to LocalDateTime and Date cells
    sheet {
        row { cell(LocalDate.of(2026, 7, 19)) } // uses dd/MM/yyyy
    }
}
```

The built-in defaults are `"yyyy-mm-dd"` and `"yyyy-mm-dd hh:mm"`. A `format` set in
the cell block, or a `format`/data format coming from a style, always wins over the
default.

## Styling

Styles come in two flavors that combine freely.

**Reusable handles** are declared once with `style { }` on the workbook builder and
applied to any number of rows and cells:

```kotlin
val header = style {
    font {
        bold = true
        color("#FFFFFF")
    }
    fill("#4472C4")
    align(horizontal = HorizontalAlignment.CENTER)
}
sheet("Report") {
    row(header) { cell("Name"); cell("Score") }
}
```

**Inline styles** are declared right on a cell with a `style { }` block for one-off
tweaks:

```kotlin
cell("Danger") {
    style { font { color(IndexedColors.RED) } }
}
```

### Fonts, fills, borders, and alignment

`font { }` sets `bold`, `italic`, `underline`, `strikeout`, `size` (points,
fractional allowed), `name`, and a color. Fills, font colors, and border colors can
come from two sources:

- **Indexed palette** — `fill(IndexedColors.LIGHT_BLUE)`, `color(IndexedColors.RED)`.
  Works in **both** XLS and XLSX.
- **Hex string** — `fill("#4472C4")`, `color("#FF0000")`. **XLSX only.**

```kotlin
val cellStyle = style {
    font { name = "Calibri"; size = 11; italic = true }
    fill("#FFF2CC")
    align(horizontal = HorizontalAlignment.RIGHT, vertical = VerticalAlignment.CENTER)
    wrapText = true
    border(BorderStyle.THIN, IndexedColors.GREY_50_PERCENT) // all four edges
    borderBottom(BorderStyle.MEDIUM)                        // override one edge
    format = "#,##0.00"
}
```

> **XLSX-only hex rule.** Hex colors are unavailable in the legacy XLS format.
> Calling `fill("#...")` or `color("#...")` does not fail immediately, but the moment
> a cell using that style is written to an `xlsWorkbook`/`xls` workbook, POI throws
> `UnsupportedOperationException` with a message pointing you at the `IndexedColors`
> overloads. In XLS, use `fill(IndexedColors.YELLOW)` / `color(IndexedColors.RED)`.

### How row, handle, and inline styles merge

When a cell is written, its final look is built by layering the styles in effect,
each layer overriding the properties the earlier ones set:

```
row style  <  cell style handle  <  inline style { }   (later wins)
```

A `format = "..."` set in the cell block overrides any data format from those
styles, and the default date/time format is applied only if nothing else set one.
Unset (`null`) properties fall through to the layer below, so styles compose rather
than replace wholesale:

```kotlin
val base = style {
    font { bold = true }
    fill("#DDDDDD")
}
sheet {
    row(base) {
        cell("plain")               // bold + grey fill from the row style
        cell("red") {
            style { font { color("#FF0000") } } // bold + grey fill + red text
        }
    }
}
```

Note that fill and font color are replaced as a unit: setting a hex fill clears an
inherited indexed fill, and vice versa.

### The dedup guarantee

No matter how you apply styles — one shared handle or a fresh inline `style { }` on
every cell — cells with an identical resulting look share one POI `CellStyle` and
`Font`. Styling a thousand cells inline still adds a single style to the workbook,
so you never have to manage or pool styles yourself to stay under Excel's limit.

## Reading cells back

Open a workbook with plain POI (e.g. `XSSFWorkbook(file)`), then use the reading
extensions on its `Sheet` and `Cell` objects. Look up cells by A1 reference with
`cellAt(...)` or the `[]` operator:

```kotlin
XSSFWorkbook(Path.of("report.xlsx").toFile()).use { wb ->
    val sheet = wb.getSheetAt(0)

    val name: String? = sheet.cellAt("A1")?.stringOrNull()
    val score: Double? = sheet["B2"]?.doubleOrNull()      // [] is an alias for cellAt
    val active: Boolean? = sheet["C2"]?.booleanOrNull()
    val joined: LocalDate? = sheet["D2"]?.localDateOrNull()
    val at: LocalDateTime? = sheet["D2"]?.localDateTimeOrNull()

    // Rendered exactly as Excel would show it, honoring the number format:
    val shown: String = sheet.cellAt("B2")?.displayString().orEmpty()
}
```

The `*OrNull` helpers return `null` when the cell is absent or holds a different
type, so they never throw on a type mismatch. For **formula cells** they inspect the
cached result type, so a formula that computes text or a number still yields its
value. `displayString()` formats any cell the way Excel would, including dates and
currencies.

## End-to-end example

Building a formatted report and reading it back:

```kotlin
import io.github.kouroshmsv.kpoi.spreadsheet.cellAt
import io.github.kouroshmsv.kpoi.spreadsheet.displayString
import io.github.kouroshmsv.kpoi.spreadsheet.xlsx
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.nio.file.Path
import java.time.LocalDate

data class Person(val name: String, val score: Double, val joined: LocalDate)

fun writeReport(target: Path) {
    val people = listOf(
        Person("Alice", 95.5, LocalDate.of(2024, 3, 12)),
        Person("Bob", 87.25, LocalDate.of(2025, 11, 2)),
        Person("Cyrus", 91.0, LocalDate.of(2026, 1, 20)),
    )

    xlsx(target) {
        val header = style {
            font {
                bold = true
                color("#FFFFFF")
            }
            fill("#4472C4")
            align(horizontal = HorizontalAlignment.CENTER)
        }
        sheet("Report") {
            freeze(rows = 1)
            columnWidth(0, characters = 18)
            row(header) {
                cell("Name")
                cell("Score")
                cell("Joined")
            }
            people.forEach { person ->
                row {
                    cell(person.name)
                    cell(person.score) { format = "0.00" }
                    cell(person.joined)                 // formatted yyyy-mm-dd
                }
            }
            row {
                cell("Total")
                formula("SUM(B2:B${people.size + 1})")
            }
            autoFilter("A1:C1")
        }
    }

    // Read it back with the kpoi helpers, as a consumer would.
    XSSFWorkbook(target.toFile()).use { wb ->
        val sheet = wb.getSheetAt(0)
        check(sheet.cellAt("A1")?.displayString() == "Name")
    }
}
```
