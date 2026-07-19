# Changelog

All notable changes to kpoi are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[Semantic Versioning](https://semver.org).

## [0.1.0] — 2026-07-19

First release.

### Added

- **kpoi-spreadsheet** — workbook / sheet / row / cell builders for XLSX,
  streaming XLSX (SXSSF), and legacy XLS; content-keyed style deduplication;
  reusable style handles and inline styles; formulas, hyperlinks, merged
  regions, freeze panes, auto-filters; date cells with configurable default
  formats; A1 read helpers (`sheet["B2"]`, `doubleOrNull()`,
  `displayString()`, …).
- **kpoi-word** — document / paragraph / run builders, headings, tables,
  inline pictures, page and line breaks, paragraph alignment and spacing.
- **kpoi-slides** — presentation / slide builders with layout placeholders,
  a title helper, free-form text boxes, bullets, pictures, and 16:9 sizing.
- **kpoi-common** — the `@PoiDsl` scope marker and hex color utilities shared
  by the other modules.

[0.1.0]: https://github.com/KouroshMsv/kpoi/releases/tag/v0.1.0
