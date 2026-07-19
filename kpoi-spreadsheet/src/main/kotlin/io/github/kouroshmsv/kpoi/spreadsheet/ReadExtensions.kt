package io.github.kouroshmsv.kpoi.spreadsheet

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellReference
import java.time.LocalDate
import java.time.LocalDateTime

/**
 * Returns the cell rendered exactly as Excel would show it, honoring the cell's
 * number format (dates, currencies, percentages, and so on). Backed by POI's
 * [org.apache.poi.ss.usermodel.DataFormatter].
 *
 * @return the formatted text; an empty string for a blank cell.
 */
public fun Cell.displayString(): String = DataFormatter().formatCellValue(this)

/**
 * Returns the text content, or `null` if the cell is not a string. For formula
 * cells the cached result type is inspected, so a formula that yields text still
 * returns its value.
 *
 * @return the string value, or `null`.
 */
public fun Cell.stringOrNull(): String? =
    if (effectiveType() == CellType.STRING) stringCellValue else null

/**
 * Returns the numeric content as a `Double`, or `null` if the cell is not
 * numeric. For formula cells the cached result type is inspected.
 *
 * @return the numeric value, or `null`.
 */
public fun Cell.doubleOrNull(): Double? =
    if (effectiveType() == CellType.NUMERIC) numericCellValue else null

/**
 * Returns the boolean content, or `null` if the cell is not boolean. For formula
 * cells the cached result type is inspected.
 *
 * @return the boolean value, or `null`.
 */
public fun Cell.booleanOrNull(): Boolean? =
    if (effectiveType() == CellType.BOOLEAN) booleanCellValue else null

/**
 * Returns the value as a [LocalDateTime], or `null` if the cell is not a
 * date-formatted numeric cell.
 *
 * @return the date-time value, or `null`.
 */
public fun Cell.localDateTimeOrNull(): LocalDateTime? =
    if (effectiveType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(this)) {
        localDateTimeCellValue
    } else {
        null
    }

/**
 * Returns the value as a [LocalDate] (the date part of [localDateTimeOrNull]), or
 * `null` if the cell is not a date-formatted numeric cell.
 *
 * @return the date value, or `null`.
 */
public fun Cell.localDateOrNull(): LocalDate? = localDateTimeOrNull()?.toLocalDate()

/**
 * Looks up a cell by A1 reference, e.g. `sheet.cellAt("B2")`.
 *
 * @param reference an A1-style cell reference.
 * @return the cell, or `null` if that row or cell does not exist.
 */
public fun Sheet.cellAt(reference: String): Cell? {
    val ref = CellReference(reference)
    return getRow(ref.row)?.getCell(ref.col.toInt())
}

/**
 * Index-operator alias for [cellAt], enabling `sheet["B2"]`.
 *
 * @param reference an A1-style cell reference.
 * @return the cell, or `null` if that row or cell does not exist.
 */
public operator fun Sheet.get(reference: String): Cell? = cellAt(reference)

private fun Cell.effectiveType(): CellType =
    if (cellType == CellType.FORMULA) cachedFormulaResultType else cellType
