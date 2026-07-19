pluginManagement {
    repositories {
        mavenCentral()
        gradlePluginPortal()
    }
}

dependencyResolutionManagement {
    repositories {
        mavenCentral()
    }
}

rootProject.name = "kpoi"

include(
    "kpoi-common",
    "kpoi-spreadsheet",
    "kpoi-word",
    "kpoi-slides",
    "samples",
)
