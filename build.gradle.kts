plugins {
    application
    kotlin("jvm") version "2.0.21"
}

group = "com.rinat"
version = "0.1.0"

repositories {
    mavenCentral()
}

dependencies {
    testImplementation(kotlin("test"))
}

tasks.test {
    useJUnitPlatform()
    maxHeapSize = "2g"
    jvmArgs("-XX:+UseG1GC")
}

kotlin {
    jvmToolchain(17)
}

application {
    mainClass.set("com.rinat.xlsxengine.MainKt")
}
