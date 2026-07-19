# Contributing to kpoi

Thanks for helping! kpoi aims to stay a thin, predictable Kotlin layer over Apache POI.

## Building

Prerequisites: JDK 17+ (21 recommended). The Gradle wrapper handles the rest.

```bash
./gradlew build          # compile + test everything
./gradlew :kpoi-spreadsheet:test
```

## Guidelines

- **Every DSL feature maps 1:1 onto POI calls.** If a feature needs cleverness beyond
  builder plumbing, discuss it in an issue first.
- **Keep the escape hatch.** New builders must expose their underlying POI object
  (`poiSomething`) as a public `val`.
- **Public API is explicit.** The project compiles with Kotlin's explicit API mode;
  new public declarations need visibility modifiers, return types, and KDoc.
- **Styles go through `StyleSpec`.** Never create `CellStyle`/`Font` objects directly in
  the spreadsheet DSL — add fields to the spec so caching keeps working.
- **Tests round-trip.** Write the document with the DSL, reopen it with plain POI, and
  assert on what a user would see.
- One focused change per PR, with a test.

## Useful references

- POI component overview: https://poi.apache.org/components/
- POI quick guides: SS (spreadsheet) busy-developers guide, XWPF and XSLF examples in the
  POI source tree under `poi-examples/`.
