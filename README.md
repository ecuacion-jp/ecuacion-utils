# ecuacion-utils

[![Java CI](https://github.com/ecuacion-jp/ecuacion-utils/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/ecuacion-jp/ecuacion-utils/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/ecuacion-jp/ecuacion-utils/branch/main/graph/badge.svg)](https://codecov.io/gh/ecuacion-jp/ecuacion-utils)
[![GitHub Release](https://img.shields.io/github/v/release/ecuacion-jp/ecuacion-utils)](https://github.com/ecuacion-jp/ecuacion-utils/releases)
[![Maven Central](https://img.shields.io/maven-central/v/jp.ecuacion.util/ecuacion-util-excel-table.svg)](https://search.maven.org/artifact/jp.ecuacion.util/ecuacion-util-excel-table)
[![Java](https://img.shields.io/badge/Java-21-ED8B00?logo=openjdk&logoColor=white)](https://www.oracle.com/java/technologies/downloads/)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)

## What is it?

`ecuacion-utils` provides Java utilities for Excel and PDF manipulation built on Apache POI and Apache PDFBox.

**What's included:**

- `ecuacion-util-excel-table` — Read/write Excel tables with header or free-format layouts (`List<List<String>>`, Bean mapping, POI Cell access)
- `ecuacion-util-excel-report-to-pdf` — Generate PDF reports from Excel templates

This library depends on `ecuacion-lib`.

## Versioning

This project follows the spirit of [Semantic Versioning](https://semver.org/). Major version increments indicate breaking changes.

## System Requirements

- JDK 21 or above.

## Documentation

(See `Documentation` part of `README` in each module)

## Installation

1. Add the required `ecuacion` modules to your `pom.xml`.
   (The following is an example for the `ecuacion-util-excel-table` module. Check the `Installation` section of the `README` in the module you want to add to your project.)

    ```xml
    <dependency>
        <groupId>jp.ecuacion.util</groupId>
        <artifactId>ecuacion-util-excel-table</artifactId>
        <!-- Put the latest release version -->
        <version>x.x.x</version>
    </dependency>

    <!-- ecuacion-lib -->
    <dependency>
        <groupId>jp.ecuacion.lib</groupId>
        <artifactId>ecuacion-lib-validation</artifactId>
        <!-- Put the latest release version -->
        <version>x.x.x</version>
    </dependency>
    ```

2. Add the required external modules to your `pom.xml`.
   (Check the `Dependent External Libraries > Manual Load Needed Libraries` section of the `README` in the module you want to add to your project.)

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for how to report bugs, suggest features, and submit pull requests.
