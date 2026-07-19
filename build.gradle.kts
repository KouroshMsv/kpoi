plugins {
    alias(libs.plugins.kotlin.jvm) apply false
    alias(libs.plugins.maven.publish) apply false
}

allprojects {
    // Maven Central namespace, verified through the KouroshMsv GitHub account.
    group = "io.github.kouroshmsv"
    version = "0.1.0"
}
