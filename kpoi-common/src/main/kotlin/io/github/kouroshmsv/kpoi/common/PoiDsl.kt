package io.github.kouroshmsv.kpoi.common

/**
 * Scope marker applied to every kpoi type-safe builder.
 *
 * `@PoiDsl` is a [DslMarker] annotation. When two receiver types in a nested
 * builder hierarchy carry the same `DslMarker`, Kotlin forbids *implicit* access
 * to an outer receiver from an inner lambda: the outer member must be called
 * explicitly (on a captured reference) or not at all. Tagging every kpoi builder
 * scope with `@PoiDsl` therefore stops mistakes such as calling `sheet { }` (a
 * member of the workbook scope) from inside a `row { }` block, where it would
 * otherwise resolve silently against the enclosing scope and build the wrong
 * structure.
 *
 * Apply it to the builder/scope classes exposed by the individual format
 * modules. It takes no parameters and has no runtime effect beyond guiding the
 * compiler.
 *
 * ```kotlin
 * // Illustration only — these scope types live in the format modules, not here.
 * @PoiDsl class SheetScope { fun row(block: RowScope.() -> Unit) { /* … */ } }
 * @PoiDsl class RowScope { fun cell(value: String) { /* … */ } }
 *
 * SheetScope().row {
 *     cell("A1")
 *     // row { }   // ← compile error: the outer SheetScope.row is not implicitly in scope
 * }
 * ```
 *
 * @see DslMarker
 */
@DslMarker
public annotation class PoiDsl
