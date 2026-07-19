package dev.kpoi.spreadsheet

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellReference
import java.time.LocalDate
import java.time.LocalDateTime

/** The cell rendered as Excel would display it, honoring number formats. */
public fun Cell.displayString(): String = DataFormatter().formatCellValue(this)

/** String content, or `null` if the cell (or its cached formula result) is not text. */
public fun Cell.stringOrNull(): String? =
    if (effectiveType() == CellType.STRING) stringCellValue else null

/** Numeric content, or `null` if the cell (or its cached formula result) is not numeric. */
public fun Cell.doubleOrNull(): Double? =
    if (effectiveType() == CellType.NUMERIC) numericCellValue else null

/** Boolean content, or `null` if the cell (or its cached formula result) is not boolean. */
public fun Cell.booleanOrNull(): Boolean? =
    if (effectiveType() == CellType.BOOLEAN) booleanCellValue else null

/** Date-time content, or `null` if the cell is not a date-formatted numeric cell. */
public fun Cell.localDateTimeOrNull(): LocalDateTime? =
    if (effectiveType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(this)) {
        localDateTimeCellValue
    } else {
        null
    }

/** Date content, or `null` if the cell is not a date-formatted numeric cell. */
public fun Cell.localDateOrNull(): LocalDate? = localDateTimeOrNull()?.toLocalDate()

/** Looks up a cell by A1 reference, e.g. `sheet.cellAt("B2")`; `null` if absent. */
public fun Sheet.cellAt(reference: String): Cell? {
    val ref = CellReference(reference)
    return getRow(ref.row)?.getCell(ref.col.toInt())
}

/** Index operator alias for [cellAt]: `sheet["B2"]`. */
public operator fun Sheet.get(reference: String): Cell? = cellAt(reference)

private fun Cell.effectiveType(): CellType =
    if (cellType == CellType.FORMULA) cachedFormulaResultType else cellType
