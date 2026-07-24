# ecuacion-util-excel-report-to-pdf

## What is it?

`ecuacion-util-excel-report-to-pdf` converts Excel (`.xlsx`) files to PDF — fonts, borders, colors, images, merged cells, headers/footers, and print settings are all faithfully reproduced.

## Usage Example

```java
PdfGenerateOptions options =
    PdfGenerateOptions.builderForExplicitFont(Path.of("/path/to/NotoSansJP-Regular.ttf"))
    .boldFontPath(Path.of("/path/to/NotoSansJP-Bold.ttf"))
    .build();

ExcelToPdfUtil.generate(
    Path.of("invoice.xlsx"),
    List.of("invoice"),          // sheet names to include
    Path.of("invoice.pdf"),
    options);
```

That's all. Point it at an Excel file, name the sheets, specify fonts, and get a PDF out.

## Dependent Ecuacion Libraries

### Automatically Loaded Libraries

(none)

### Manual Load Needed Libraries

- `ecuacion-lib-core`

## Dependent External Libraries

### Automatically Loaded Libraries

- `org.apache.pdfbox`
- `org.apache.poi:poi`
- `org.apache.poi:poi-ooxml`
- `jakarta.validation:jakarta.validation-api`
- `jakarta.mail:jakarta.mail-api`
- `org.slf4j:slf4j-api`
- `org.apache.commons:commons-lang3`

### Manual Load Needed Libraries

- (any `jakarta.validation:jakarta.validation-api` compatible Jakarta Validation implementation. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- (any `org.slf4j:slf4j-api` compatible logging implementation. `ch.qos.logback:logback-classic` is recommended.)

## Documentation

- [ecuacion-references-utils](https://references.ecuacion.jp/ecuacion-references-utils/public/showMarkdown/page?id=home) — Official reference documentation
- [javadoc](https://javadoc.io/doc/jp.ecuacion.util/ecuacion-util-excel-report-to-pdf/latest/jp.ecuacion.util.pdf.excel.report/module-summary.html)

## Sample Code

- [ecuacion-util-excel-report-to-pdf-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-excel-report-to-pdf-sample)

## Installation

Check [Installation](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
The description of dependent `ecuacion` modules is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-excel-report-to-pdf</artifactId>
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

