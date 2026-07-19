# Module kpoi-common

`kpoi-common` is the shared-infrastructure module of **kpoi**, a small Kotlin DSL
over [Apache POI](https://poi.apache.org/). It carries the handful of types that
belong to no single POI format yet are needed by all of them, so the per-format
modules (spreadsheets, word-processing documents, presentations) can depend on one
common vocabulary instead of each re-inventing it. Everything here lives in the
package `io.github.kouroshmsv.kpoi.common`.

The module is deliberately tiny and exposes just two public declarations. The
first, `PoiDsl`, is the `@DslMarker` annotation that every kpoi builder is tagged
with, so the nested type-safe builders (`workbook { sheet { row { … } } }` and
their relatives) enforce correct scoping at compile time. The second, `Rgb`, is a
validated sRGB color value that bridges plain hex strings to the three different
color representations the various POI formats expect.

The mental model is **"one shared color, three destinations"**: you describe a
color once — usually by calling `Rgb.parse("#RRGGBB")` — and each downstream module
asks that value for the shape it needs. XSSF (`.xlsx`) wants a raw byte array, XWPF
(`.docx`) wants a hex string, and XSLF (`.pptx`) wants a `java.awt.Color`. Keeping
that conversion in one place means the higher-level DSLs never have to reason about
POI's per-format color quirks, and the `@PoiDsl` marker keeps their builder blocks
honest.

Because the other kpoi modules depend on this one, treat the two types below as a
stable, shared contract rather than internal helpers.

# Package io.github.kouroshmsv.kpoi.common

This package holds the DSL scope marker and the color utility used across kpoi.

## `@PoiDsl` — the DSL scope marker

`PoiDsl` is an annotation meta-annotated with Kotlin's
[`@DslMarker`](https://kotlinlang.org/docs/type-safe-builders.html#scope-control-dslmarker).
A `DslMarker` controls implicit-receiver resolution inside nested builder lambdas:
when two receiver types in a builder hierarchy carry the *same* marker, Kotlin
refuses to implicitly resolve a call against an outer receiver from within an inner
lambda. The outer member must then be invoked explicitly on a captured reference,
or the code simply will not compile.

Tagging every kpoi builder scope with `@PoiDsl` prevents a whole class of quiet
bugs — for example calling `sheet { }` (a member of the workbook scope) from inside
a `row { }` block, which without scope control would resolve against the enclosing
workbook and build the wrong structure. The annotation takes no parameters and has
no runtime behavior of its own; it exists purely to guide the compiler.

You apply it to the builder/scope classes exposed by the individual format modules:

```kotlin
// Illustration only — the concrete scope types live in the format modules.
@PoiDsl class SheetScope { fun row(block: RowScope.() -> Unit) { /* … */ } }
@PoiDsl class RowScope { fun cell(value: String) { /* … */ } }

SheetScope().row {
    cell("A1")
    // row { }   // ← compile error: outer SheetScope.row is not implicitly in scope
}
```

## `Rgb` — a validated sRGB color

`Rgb` is a `data class` holding three integer channels, `red`, `green` and `blue`,
each constrained to `0..255`. The constructor validates its arguments: passing a
channel outside that range throws `IllegalArgumentException` (the message reports
the offending triple, e.g. `Color channels must be in 0..255, got (256, 0, 0)`).
Being a `data class`, it gains structural `equals`/`hashCode`, a readable
`toString`, and `copy` for free.

```kotlin
val red = Rgb(255, 0, 0)         // ok
// val bad = Rgb(256, 0, 0)      // throws IllegalArgumentException
```

### Creating an `Rgb` from hex — `Rgb.parse`

Most callers do not build channels by hand; they parse a hex string with the
companion factory `Rgb.parse`. It accepts both `"#RRGGBB"` and `"RRGGBB"` (a single
leading `#` is optional) and is case-insensitive. After any `#` is removed, the
remaining text must be exactly six hexadecimal digits (`0`–`9`, `a`–`f`); anything
else throws `IllegalArgumentException`.

```kotlin
Rgb.parse("#4472C4")   // Rgb(red = 68, green = 114, blue = 196)
Rgb.parse("4472c4")    // same value — the '#' is optional and case is ignored

Rgb.parse("#12345")    // throws IllegalArgumentException (too few digits)
Rgb.parse("blue")      // throws IllegalArgumentException (not hex)
Rgb.parse("#GGGGGG")   // throws IllegalArgumentException (G is not a hex digit)
```

### Converting to what POI expects

Once you hold an `Rgb`, three zero-argument conversions produce the format-specific
representations. Each is a pure function returning a fresh value:

| Method | Returns | Consumed by |
| --- | --- | --- |
| `toAwtColor()` | `java.awt.Color` | XSLF (`.pptx` presentations) |
| `toByteArray()` | 3-element `ByteArray`, `[r, g, b]` | XSSF (`.xlsx` spreadsheets) |
| `toHexString()` | uppercase `"RRGGBB"` string, no `#` | XWPF (`.docx` documents) |

```kotlin
val orange = Rgb.parse("#FF8000")     // Rgb(red = 255, green = 128, blue = 0)

orange.toAwtColor()                   // java.awt.Color(255, 128, 0)
orange.toHexString()                  // "FF8000"
orange.toByteArray().toList()         // [-1, -128, 0]
```

Note the byte values in the last line. Kotlin's `Byte` is signed, so any channel
above `127` shows up as a negative number — `255` becomes `-1`, `128` becomes
`-128` — even though the underlying 8-bit pattern is exactly what POI wants.
`toHexString` always yields six upper-case characters because each channel is
capped at `255` and formatted with a two-digit `%02X`.

### Round-tripping

`parse` and `toHexString` are inverses up to letter case, which makes `Rgb` a
convenient normalization point for user-supplied colors:

```kotlin
Rgb.parse("#4472c4").toHexString()    // "4472C4"
```
