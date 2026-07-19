import com.vanniktech.maven.publish.JavadocJar
import com.vanniktech.maven.publish.KotlinJvm
import org.gradle.plugins.signing.SigningExtension
import org.jetbrains.kotlin.gradle.dsl.JvmTarget

plugins {
    alias(libs.plugins.kotlin.jvm)
    alias(libs.plugins.maven.publish)
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
}

dependencies {
    testImplementation(libs.junit.jupiter)
    testRuntimeOnly(libs.junit.launcher)
}

tasks.test {
    useJUnitPlatform()
}

mavenPublishing {
    publishToMavenCentral()
    configure(KotlinJvm(javadocJar = JavadocJar.Empty(), sourcesJar = true))
    // Sign only when a key is configured, so contributors can build without one.
    if (providers.gradleProperty("signing.gnupg.keyName").isPresent ||
        providers.gradleProperty("signingInMemoryKey").isPresent
    ) {
        signAllPublications()
    }
    pom {
        name = "kpoi-common"
        description = "Shared DSL infrastructure for kpoi, a Kotlin DSL for Apache POI"
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

if (providers.gradleProperty("signing.gnupg.keyName").isPresent) {
    configure<SigningExtension> {
        useGpgCmd()
    }
}
