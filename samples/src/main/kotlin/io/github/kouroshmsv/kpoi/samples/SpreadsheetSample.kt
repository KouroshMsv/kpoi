package io.github.kouroshmsv.kpoi.samples

import io.github.kouroshmsv.kpoi.spreadsheet.cellAt
import io.github.kouroshmsv.kpoi.spreadsheet.displayString
import io.github.kouroshmsv.kpoi.spreadsheet.xlsx
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.nio.file.Path
import java.time.LocalDate

data class Person(val name: String, val score: Double, val joined: LocalDate)

fun spreadsheetSample(outputDir: Path): Path {
    val people = listOf(
        Person("Alice", 95.5, LocalDate.of(2024, 3, 12)),
        Person("Bob", 87.25, LocalDate.of(2025, 11, 2)),
        Person("Cyrus", 91.0, LocalDate.of(2026, 1, 20)),
    )

    val target = outputDir.resolve("report.xlsx")
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
                    cell(person.joined)
                }
            }
            row {
                cell("Total")
                formula("SUM(B2:B${people.size + 1})")
            }
            autoFilter("A1:C1")
        }
    }

    // Read it back with the kpoi read helpers, as a consumer would.
    XSSFWorkbook(target.toFile()).use { workbook ->
        val sheet = workbook.getSheetAt(0)
        check(sheet.cellAt("A1")?.displayString() == "Name")
    }
    return target
}
