# ecuacion-utils

[![Java](https://img.shields.io/badge/Java-21-ED8B00?logo=openjdk&logoColor=white)](https://www.oracle.com/java/technologies/downloads/)
[![Java CI](https://github.com/ecuacion-jp/ecuacion-utils/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/ecuacion-jp/ecuacion-utils/actions/workflows/ci.yml)
[![codecov](https://codecov.io/gh/ecuacion-jp/ecuacion-utils/branch/main/graph/badge.svg)](https://codecov.io/gh/ecuacion-jp/ecuacion-utils)
[![GitHub Release](https://img.shields.io/github/v/release/ecuacion-jp/ecuacion-utils)](https://github.com/ecuacion-jp/ecuacion-utils/releases)
[![Maven Central](https://img.shields.io/maven-central/v/jp.ecuacion.util/ecuacion-util-excel-table.svg)](https://search.maven.org/artifact/jp.ecuacion.util/ecuacion-util-excel-table)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://www.apache.org/licenses/LICENSE-2.0)

## What is it?

`ecuacion-utils` provides Java utilities for Excel and PDF manipulation built on Apache POI and Apache PDFBox.

**What's included:**

- `ecuacion-util-excel-table` — Read/write Excel tables with header or free-format layouts (`List<List<String>>`, Bean mapping, POI Cell access)
- `ecuacion-util-excel-report-to-pdf` — Generate PDF reports from Excel templates

This library depends on `ecuacion-lib`. Major versions are compatible as follows:

| ecuacion-utils | ecuacion-lib |
| --- | --- |
| 5.x | 16.x |

## Supported Languages

Messages (error messages of Excel and PDF operations) are provided in the following languages, the same as `ecuacion-lib`.
The language is chosen from the `Locale` passed to `ecuacion-lib`.

| Language | Locale | Language | Locale |
| --- | --- | --- | --- |
| English (default) | `en` | Japanese | `ja` |
| Arabic | `ar` | Azerbaijani | `az` |
| Bengali | `bn` | Chinese (Simplified) | `zh` |
| Chinese (Traditional) | `zh_Hant` (also used for `zh_TW`, `zh_HK`, `zh_MO`) | Czech | `cs` |
| Danish | `da` | Dutch | `nl` |
| Filipino | `fil` | French | `fr` |
| German | `de` | Gujarati | `gu` |
| Hindi | `hi` | Hungarian | `hu` |
| Indonesian | `id` | Italian | `it` |
| Kannada | `kn` | Korean | `ko` |
| Malay | `ms` | Malayalam | `ml` |
| Marathi | `mr` | Mongolian | `mn` |
| Persian | `fa` | Polish | `pl` |
| Portuguese | `pt` | Romanian | `ro` |
| Russian | `ru` | Slovak | `sk` |
| Spanish | `es` | Tamil | `ta` |
| Telugu | `te` | Thai | `th` |
| Turkish | `tr` | Ukrainian | `uk` |
| Vietnamese | `vi` | | |

Translations other than English and Japanese are based on machine translation and may contain unnatural expressions.
Corrections and new languages are welcome via pull requests. See [Contributing](#contributing).

## Versioning

This project follows the spirit of [Semantic Versioning](https://semver.org/). Major version increments indicate breaking changes.

## System Requirements

- JDK 21 or above.

## Documentation

- Official reference documentation - [ecuacion-references-utils](https://references.ecuacion.jp/ecuacion-references-utils/public/showMarkdown/page?id=home)
- javadoc - [javadoc.io](https://javadoc.io/doc/jp.ecuacion.util/)

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

    > **Note:** `ecuacion-lib-validation` is listed here rather than the directly-required
    > `ecuacion-lib-core`, so that `ecuacion-lib-validation`'s custom validators are usable
    > out of the box with no extra setup. `ecuacion-lib-core` is included transitively
    > (`ecuacion-lib-validation` depends on it).

2. Add the required external modules to your `pom.xml`.
   (Check the `Dependent External Libraries > Manual Load Needed Libraries` section of the `README` in the module you want to add to your project.)

## Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for how to report bugs, suggest features, and submit pull requests.
