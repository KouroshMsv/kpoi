import org.jetbrains.kotlin.gradle.dsl.JvmTarget

// Runnable examples; compiled by CI so the README snippets can never rot.
// This module is not published.
plugins {
    alias(libs.plugins.kotlin.jvm)
    application
}

kotlin {
    compilerOptions {
        jvmTarget = JvmTarget.JVM_11
    }
}

java {
    sourceCompatibility = JavaVersion.VERSION_11
    targetCompatibility = JavaVersion.VERSION_11
}

dependencies {
    implementation(project(":kpoi-spreadsheet"))
    implementation(project(":kpoi-word"))
    implementation(project(":kpoi-slides"))
    runtimeOnly(libs.log4j.core)
}

application {
    mainClass = "io.github.kouroshmsv.kpoi.samples.MainKt"
}
