plugins {
    alias(libs.plugins.kotlin.jvm) apply false
}

allprojects {
    // Maven Central namespace, verified through the KouroshMsv GitHub account.
    group = "io.github.kouroshmsv"
    version = "0.1.0-SNAPSHOT"
}
