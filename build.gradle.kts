plugins {
    alias(libs.plugins.kotlin.jvm) apply false
    alias(libs.plugins.maven.publish) apply false
    alias(libs.plugins.dokka)
}

allprojects {
    // Maven Central namespace, verified through the KouroshMsv GitHub account.
    group = "io.github.kouroshmsv"
    version = "0.2.0-SNAPSHOT"
}

// Aggregated API docs: ./gradlew :dokkaGenerate -> build/dokka/html
dependencies {
    dokka(project(":kpoi-common"))
    dokka(project(":kpoi-spreadsheet"))
    dokka(project(":kpoi-word"))
    dokka(project(":kpoi-slides"))
}
