package io.github.kouroshmsv.kpoi.samples

import java.nio.file.Files
import java.nio.file.Path

fun main() {
    val outputDir = Path.of("build", "sample-output")
    Files.createDirectories(outputDir)
    println("spreadsheet: " + spreadsheetSample(outputDir).toAbsolutePath())
    println("word:        " + wordSample(outputDir).toAbsolutePath())
    println("slides:      " + slidesSample(outputDir).toAbsolutePath())
}
