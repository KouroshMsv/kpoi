package dev.kpoi.common

/**
 * Scope marker for all kpoi builders.
 *
 * Prevents accidentally calling a method of an outer builder from a nested
 * builder lambda (e.g. calling `sheet {}` from inside a `row {}` block).
 */
@DslMarker
public annotation class PoiDsl
