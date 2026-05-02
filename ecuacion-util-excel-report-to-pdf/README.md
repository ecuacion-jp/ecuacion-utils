# ecuacion-util-excel-report-to-pdf

## What is it?

`ecuacion-util-excel-report-to-pdf` provides utilities for `apache PDFBox`.  

## System Requirements

- JDK 21 or above.

## Dependent Ecuacion Libraries

### Automatically Loaded Libraries

(none)

### Manual Load Needed Libraries

- `ecuacion-lib-core`

## Dependent External Libraries

### Automatically Loaded External Libraries

- `org.apache.pdfbox`
- `org.apache.poi:poi`
- `org.apache.poi:poi-ooxml`

### Manually Load Needed External Libraries

- `jakarta.validation:jakarta.validation-api`
- (any `jakarta.validation:jakarta.validation-api` compatible Bean Validation libraries. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- `jakarta.annotation:jakarta.annotation-api`
- `org.slf4j:slf4j-api`
- (if you use log4j2, add `org.apache.logging.log4j:log4j-slf4j-impl` and `org.apache.logging.log4j:log4j-core`,
   or else `org.apache.logging.log4j.log4j-to-slf4j` (To use any slf4j-compatible logging modules) and any `org.slf4j:slf4j-api` compatible logging libraries. `ch.qos.logback:logback-classic` is recommended.)

(modules depending on `ecuacion-lib-core`)

- `jakarta.mail:jakarta.mail-api` (If you want to use the mail related utility: `jp.ecuacion.lib.core.util.MailUtil`)

Since the dependency libraries are a little complicated, we recommend to refer `pom.xml` in `ecuacion-util-excel-report-to-pdf-sample`.

## Documentation

- [javadoc](https://javadoc.ecuacion.jp/apidocs/ecuacion-util-excel-report-to-pdf/)

## Sample Code

- [ecuacion-util-excel-report-to-pdf-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-excel-report-to-pdf-sample)

## Introduction

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
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-lib-core</artifactId>
    <!-- Put the latest release version -->
    <version>x.x.x</version>
</dependency>
```

## Features

### Create PDF from Excel

#### Confirmed Supported Features

The following features have been verified through automated tests.

- **Page size and orientation**
  - Supported paper sizes: A4, A5, Letter
  - Portrait and landscape both supported
- **Margins (left and top)**
  - Left and top margin values set in Excel are accurately reflected in the PDF
  - Verified with Normal (left=0.7in), Wide (left=1.0in), Narrow (left=0.25in), and custom values
- **Print area**
  - Only cells within the explicitly specified print area are rendered in the PDF
  - When no print area is set, the used cell range is printed
- **Page breaks**
  - Manual row breaks split content across multiple pages at the specified row
  - Manual column breaks split content across multiple pages at the specified column
- **Scale (explicit percentage)**
  - Explicit scale percentage (e.g., 50%, 100%, 200%) is accurately applied to cell sizes in the PDF
- **Fit to page (no explicit scale)**
  - Content wider or taller than the printable area is automatically scaled down to fit one page
  - Content that fits naturally is rendered at its natural size without scaling
- **Header and footer**
  - Header and footer text is rendered within the header/footer margin area configured in Excel
  - Supported format codes: `&L`/`&C`/`&R` (section alignment), `&P`/`&P+n`/`&P-n` (page
    number), `&N` (total pages), `&D` (date), `&T` (time), `&A` (sheet name), `&F` (file name),
    `&Z` (file path), `&B` (bold), `&U` (underline), `&E` (double underline), `&S`
    (strikethrough), `&X`/`&Y` (superscript/subscript), `&nn` (font size), `&"Font,Style"`
    (font), `&KRRGGBB` (color), `&&` (literal &)
  - `&I` (italic) is parsed but has no effect as the embedded font has no italic face
  - **Images in headers/footers (`&G`) are not supported.** Place images directly in sheet
    cells instead.
- **Cell borders**
  - Solid borders: THIN, MEDIUM, THICK are rendered at their respective line widths (0.5 pt, 1.0 pt, 1.5 pt)
  - Dashed/dotted borders: DASHED, DOTTED, DASH_DOT, DASH_DOT_DOT, MEDIUM_DASHED,
    MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT, SLANTED_DASH_DOT are rendered with the
    corresponding dash pattern
  - Border color is reflected in the PDF
  - All four sides (top, bottom, left, right) are independently supported
  - **Diagonal borders**: both ↘ (top-left to bottom-right) and ↗ (bottom-left to top-right)
    are supported, including dashed styles; both can be active simultaneously
  - DOUBLE border style is not supported
- **Cell text**
  - Font size is accurately reflected in the PDF
  - **Bold** is supported
  - **Italic** is supported (rendered as synthetic italic using a shear transformation;
    the embedded font has no true italic face)
  - **Strikethrough** is supported
  - **Superscript** and **Subscript** are supported (rendered at 70% size with adjusted
    baseline)
  - Text color is reflected in the PDF
  - **Horizontal alignment**: LEFT, CENTER, RIGHT are supported; GENERAL alignment
    right-aligns numeric and date values and left-aligns strings, matching Excel's default
    behavior
  - **Vertical alignment**: TOP, CENTER, BOTTOM are supported
  - Formula cells display their cached result value, not the formula string; alignment
    follows the result type (numeric → right, string → left under GENERAL alignment)
  - **Text wrapping** (`wrapText`) is supported; long text is wrapped to multiple lines
    within the cell
  - **Shrink to fit** (`shrinkToFit`) is supported; font size is reduced so that
    single-line text fits within the cell width
  - **Vertical text** (rotation = 255, 縦書き) is supported; characters are stacked
    top to bottom and centered horizontally within the cell
  - Text that overflows the cell height is clipped: with TOP alignment lower lines are
    hidden; with BOTTOM or CENTER alignment upper lines are hidden
  - **Underline** (`U_SINGLE`, `U_SINGLE_ACCOUNTING`) and **double underline**
    (`U_DOUBLE`, `U_DOUBLE_ACCOUNTING`) are supported; accounting variants span the
    full cell width while standard variants span only the text width
  - **Angle rotation** (values 1–254) is not supported; only vertical text
    (rotation = 255) is rendered
  - Text overflow into adjacent empty cells is not reproduced; text that exceeds the
    cell width may be visible beyond the cell boundary
- **Cell number and date formats**
  - Standard number formats are rendered correctly: integer (`0`), decimal (`0.00`),
    thousands separator (`#,##0`), percentage (`0%`, `0.0%`), currency (`¥#,##0`,
    `$#,##0.00`), scientific notation (`0.00E+00`), and negative numbers
  - Date formats: `yyyy/mm/dd`, `mm/dd/yyyy`, `dd/mm/yyyy`, `yyyy-mm-dd`,
    `yyyy年m月d日`, and two-digit year (`yy/mm/dd`) are supported
  - Abbreviated and full month names (`mmm`, `mmmm`) are rendered in the locale
    specified by the `[$-xxx]` code in the format string (defaults to English)
  - Japanese weekday names via `aaa`/`aaaa` tokens (e.g., `月`, `月曜日`) and
    English weekday names via `ddd`/`dddd` tokens (e.g., `Wed`, `Wednesday`)
  - Time-only formats (`h:mm`, `h:mm:ss`, `h:mm AM/PM`) and date-time combined
    formats (e.g., `yyyy/mm/dd h:mm:ss`) are supported
  - Formula cells display their cached result value formatted according to the
    cell's format string
  - **Japanese era (和暦)**: the Reiwa era (`ggge"年"m"月"d"日"` and similar) is
    supported for dates from 2019-05-01 onward. Dates before 2019-05-01 throw
    a `RuntimeException` at PDF generation time. Other eras (Heisei and earlier)
    are not supported.
- **Cell background color**
  - Solid fill (`SOLID_FOREGROUND`) is rendered with the exact RGB color value,
    including theme colors with tint applied (lighter/darker variations)
  - Cells with no fill are rendered with a white background
  - Indexed colors (legacy palette) are supported
  - Theme colors are supported; tint values (positive = lighter, negative = darker)
    are correctly applied via `getRGBWithTint()`
  - **Non-solid fill patterns** (e.g., `THIN_HORZ_BANDS`, `BRICKS`, etc.) are not
    rendered as background fills — only `SOLID_FOREGROUND` is supported
  - Background fill is drawn before text, so text is rendered on top of the
    background color as expected
- **Images**
  - PNG and JPEG images are supported; transparent PNG images retain their
    transparency when rendered in the PDF
  - Both two-cell anchor (resizes with rows/columns) and one-cell anchor
    (fixed position relative to top-left cell) are supported
  - Multiple images on a single sheet are all rendered
  - Images on multiple pages are each rendered on the correct page
  - Image scale is applied correctly
- **Merged cells**
  - Horizontally, vertically, and rectangular merged regions are all supported
  - The border of a merged region uses the style of the outermost cells: right
    border from the rightmost cell, bottom border from the bottom-row cell
  - Background fill is applied across the entire merged region
  - Text alignment (horizontal and vertical) within a merged region is respected
- **Print title rows**
  - When "Rows to repeat at top" is configured in Excel's Page Setup, those rows
    are rendered at the top of every page of the PDF
- **Print title columns**
  - When "Columns to repeat at left" is configured in Excel's Page Setup, those
    columns are rendered at the left of every page of the PDF
  - Can be combined with print title rows; both are rendered on every page
- **Multiple sheets**
  - Multiple sheets can be specified for a single PDF export; each sheet's pages
    are appended in order to produce a single PDF document

#### Excel Values used to create PDF files

- File > Page Setup > Page tab > orientation
- File > Page Setup > Page tab > margins
