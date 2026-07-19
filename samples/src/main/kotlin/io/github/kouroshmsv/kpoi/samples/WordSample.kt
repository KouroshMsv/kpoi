package io.github.kouroshmsv.kpoi.samples

import io.github.kouroshmsv.kpoi.word.docx
import org.apache.poi.xwpf.usermodel.ParagraphAlignment
import java.nio.file.Path

fun wordSample(outputDir: Path): Path {
    val target = outputDir.resolve("letter.docx")
    docx(target) {
        heading("Quarterly Letter", level = 1)
        paragraph {
            text("Dear ")
            text("shareholders") { bold = true }
            text(",")
        }
        paragraph("Results were strong across all three regions.") {
            alignment = ParagraphAlignment.BOTH
            spacingAfterPoints = 12
        }
        table(width = "100%") {
            row {
                cell("Metric")
                cell("Value")
            }
            row {
                cell("Revenue")
                cell("$12M")
            }
            row {
                cell("NPS")
                cell("61")
            }
        }
        pageBreak()
        paragraph("Appendix follows.")
    }
    return target
}
