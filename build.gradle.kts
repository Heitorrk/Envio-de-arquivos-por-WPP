// build.gradle.kts
plugins {
    application
    id("com.github.johnrengelman.shadow") version "8.1.1"
}

java {
    toolchain { languageVersion.set(org.gradle.jvm.toolchain.JavaLanguageVersion.of(21)) }
}

repositories { mavenCentral() }

dependencies {
    // Selenium 4 (Selenium Manager baixa o ChromeDriver automaticamente)
    implementation("org.seleniumhq.selenium:selenium-java:4.23.0")

    // Apache POI para Excel (XSSFWorkbook, DataFormatter, etc.)
    implementation("org.apache.poi:poi-ooxml:5.2.5")

    // PDFBox 3.x (compatível com org.apache.pdfbox.Loader)
    implementation("org.apache.pdfbox:pdfbox:3.0.2")

    // (Opcional) utilidades
    // implementation("org.json:json:20240303")
    // implementation("commons-io:commons-io:2.16.1")
}

application {
    // Sua Main confirmada
    mainClass.set("org.contraApp.Main")
}

tasks.withType<JavaCompile> {
    options.release.set(21)
}

tasks.withType<Jar> {
    // Evita conflitos de arquivos META-INF ao empacotar tudo
    duplicatesStrategy = DuplicatesStrategy.EXCLUDE
}

tasks.named<com.github.jengelman.gradle.plugins.shadow.tasks.ShadowJar>("shadowJar") {
    archiveBaseName.set("evox-contrax")
    archiveClassifier.set("all")
    archiveVersion.set("")
    manifest { attributes["Main-Class"] = application.mainClass.get() }

    // Se quiser reduzir o jar, habilite o minimize() (cuidado com libs que usam reflexão)
    // minimize()
}

tasks.build { dependsOn(tasks.shadowJar) }
