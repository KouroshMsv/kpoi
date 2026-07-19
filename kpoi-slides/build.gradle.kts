import org.jetbrains.kotlin.gradle.dsl.JvmTarget

plugins {
    alias(libs.plugins.kotlin.jvm)
    `maven-publish`
}

kotlin {
    explicitApi()
    compilerOptions {
        jvmTarget = JvmTarget.JVM_11
    }
}

java {
    sourceCompatibility = JavaVersion.VERSION_11
    targetCompatibility = JavaVersion.VERSION_11
    withSourcesJar()
}

dependencies {
    api(project(":kpoi-common"))
    api(libs.poi.ooxml)

    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.launcher)
    testRuntimeOnly(libs.log4j.core)
}

tasks.test {
    useJUnitPlatform()
}

publishing {
    publications {
        create<MavenPublication>("maven") {
            from(components["java"])
            pom {
                name = "kpoi-slides"
                description = "Kotlin DSL for Apache POI PowerPoint presentations (XSLF)"
                url = "https://github.com/KouroshMsv/kpoi"
                licenses {
                    license {
                        name = "The Apache License, Version 2.0"
                        url = "https://www.apache.org/licenses/LICENSE-2.0.txt"
                    }
                }
                developers {
                    developer {
                        id = "KouroshMsv"
                        name = "Kourosh"
                    }
                }
                scm {
                    url = "https://github.com/KouroshMsv/kpoi"
                    connection = "scm:git:https://github.com/KouroshMsv/kpoi.git"
                    developerConnection = "scm:git:git@github.com:KouroshMsv/kpoi.git"
                }
            }
        }
    }
}
