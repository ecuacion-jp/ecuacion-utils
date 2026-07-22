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
  - Explicit scale percentage (e.g., 50%, 100%, 200%) is accurately applied to cell geometry (row heights, column widths) and font sizes in the PDF
- **Fit to page (no explicit scale)**
  - Content wider or taller than the printable area is automatically scaled down to fit one page
  - Content that fits naturally is rendered at its natural size without scaling
  - Both column-width constraint and row-height constraint are applied: the tighter of the two determines the final scale factor, ensuring all content fits on one page in both dimensions
  - The scale factor is applied uniformly to row heights, column widths, **and font sizes**, matching Excel's cm-matrix scaling behaviour
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
  - **Indentation**: cell indent levels are applied; LEFT-aligned text is shifted right,
    RIGHT-aligned text is shifted left, by approximately one character width per indent level
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
  - Text is rendered correctly even when the font line height slightly exceeds the cell
    height (e.g., large bold font in a tight row after fit-to-page scaling); previously
    a floating-point guard could silently suppress all text in such cells
- **Cell number and date formats (additional)**
  - Accounting-style zero section with quoted literal and `??` alignment (e.g., `"-"??`)
    renders as the quoted literal `"-"` for zero-valued formula cells; previously the `??`
    placeholder was incorrectly rendered as the digit `0`, producing `"- 0"`
  - **Underline** (`U_SINGLE`, `U_SINGLE_ACCOUNTING`) and **double underline**
    (`U_DOUBLE`, `U_DOUBLE_ACCOUNTING`) are supported; accounting variants span the
    full cell width while standard variants span only the text width
  - **Angle rotation** (values 1–254) is not supported; only vertical text
    (rotation = 255) is rendered
  - Numeric values with GENERAL alignment are correctly right-aligned within the cell
    boundary, regardless of whether adjacent cells are empty
  - String text that exceeds the cell width is clipped at the cell boundary; Excel's
    behavior of visually extending long strings into adjacent empty cells is not reproduced
- **Cell number and date formats**
  - Standard number formats are rendered correctly: integer (`0`), decimal (`0.00`),
    thousands separator (`#,##0`), percentage (`0%`, `0.0%`), currency (`¥#,##0`,
    `$#,##0.00`), scientific notation (`0.00E+00`), and negative numbers
  - Date formats: `yyyy/mm/dd`, `mm/dd/yyyy`, `dd/mm/yyyy`, `yyyy-mm-dd`,
    `yyyy年m月d日`, and two-digit year (`yy/mm/dd`) are supported
  - Locale-sensitive built-in date formats (e.g., Excel format ID 14) are rendered
    using `DateTimeFormatter.ofLocalizedDate` with the locale from
    `PdfGenerateOptions.dateLocale()`, falling back to `Locale.getDefault()` when not
    set; this covers all locales automatically (e.g., Japanese → `2018/02/21`,
    US → `2/21/18`, Italian → `21/02/18`)
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
  - Images and auto-shapes anchored outside the print area columns or rows are
    excluded from the PDF output, matching Excel's print behavior
- **Merged cells**
  - Horizontally, vertically, and rectangular merged regions are all supported
  - The border of a merged region uses the style of the outermost cells: right
    border from the rightmost cell, bottom border from the bottom-row cell
  - Background fill is applied across the entire merged region
  - Text alignment (horizontal and vertical) within a merged region is respected
- **Print title rows**
  - When "Rows to repeat at top" is configured in Excel's Page Setup, those rows
    are rendered at the top of every page of the PDF
  - Rows that appear before the print title row (preamble rows) are rendered on
    the first page only, above the repeated title row, matching Excel's behavior
- **Print title columns**
  - When "Columns to repeat at left" is configured in Excel's Page Setup, those
    columns are rendered at the left of every page of the PDF
  - Can be combined with print title rows; both are rendered on every page
- **Excel table styles**
  - When a cell range is defined as an Excel table (`Insert > Table`) with a table style,
    the style is applied to the PDF output
  - Header row fill, alternating row stripes (first/second row stripe), and first/last
    column highlights are all rendered from the table style
  - Internal horizontal borders (row separators) and outer table borders defined in the
    table style's `wholeTable` element are rendered
  - Font colour overrides for the header row are applied; when the header fill is dark
    and no explicit font colour is resolvable, white text is used automatically
  - Both built-in Excel table styles (e.g. `TableStyleMedium6`) and custom table styles
    defined within the workbook are supported
- **Multiple sheets**
  - Multiple sheets can be specified for a single PDF export; each sheet's pages
    are appended in order to produce a single PDF document

### System font support (`useSystemFonts`)

By default, font files must be specified explicitly via `PdfGenerateOptions.builderForExplicitFont`.
Using `PdfGenerateOptions.builderForSystemFonts()` instead, the library automatically searches the
OS font directories for the font that matches the workbook's default font. The located font is
embedded in the output PDF and is used for accurate column-width calculation.

```java
PdfGenerateOptions options = PdfGenerateOptions.builderForSystemFonts()
    .build();   // regularFontPath is optional here — it acts as a fallback
```

When `builderForSystemFonts()` is used and no matching system font is found, a
`PdfGenerateException` is thrown. Specifying `regularFontPath` in addition acts as a fallback for
that case.

When a font family ships multiple weight variants (e.g. 游ゴシック Light / Medium / Regular),
the library prefers the **Medium** weight over Light for the regular (non-bold) font, matching
Excel on macOS which uses Medium as the default display weight for CJK fonts.
The font selection also correctly excludes **italic** and **bold** variants when a regular (upright)
font is requested, so fonts such as Calibri are reliably resolved to their Regular face.

When the workbook's default font does not contain CJK glyphs (e.g. Calibri), any CJK characters
in cell text are automatically rendered using the fallback font specified by `regularFontPath`,
on a character-by-character basis.

#### Multiple fallback fonts

`regularFontPath`/`boldFontPath` act as the first fallback font tried when a character can't be
encoded by the primary font. For workbooks mixing three or more scripts where no single fallback
font covers all of them (e.g. a report mixing Japanese, Korean, and Arabic text), additional
fallback fonts can be registered via `addFallbackFont`, which can be called multiple times; each
character is tried against the primary font, then each fallback in registration order, until one
can encode it:

```java
PdfGenerateOptions options = PdfGenerateOptions.builderForSystemFonts()
    .regularFontPath(japaneseFontPath)         // first fallback
    .addFallbackFont(koreanFontPath, null)     // second fallback
    .addFallbackFont(arabicFontPath, null)     // third fallback
    .build();
```

This also applies in explicit font mode (`builderForExplicitFont`): `addFallbackFont` fonts are
tried for characters the primary `regularFontPath` font cannot encode.

#### Per-cell font resolution

When a cell's font differs from the workbook's default font (e.g. most of the report uses
Calibri, but a few cells are explicitly set to a Japanese font for localized content), that
cell's own font is resolved from the OS font directories and used for the cell, instead of
always using the workbook's default font. This applies to each distinct font name found in the
workbook, following the same weight-variant and TTC lookup rules described above.

If a cell's font cannot be found on the system, the cell falls back to the workbook's default
font (a warning is logged) rather than failing PDF generation — generation only fails when the
workbook's default font itself cannot be resolved and no `regularFontPath` fallback is set.

> **Font licensing notice:** the located system font is embedded in the output PDF.
> Confirm that the font's licence permits embedding and distribution before enabling this option.

### PDF password protection

Setting `pdfPassword` in `PdfGenerateOptions` encrypts the output PDF with **256-bit AES**.
The PDF can only be opened with the specified password.

```java
PdfGenerateOptions options =
    PdfGenerateOptions.builderForExplicitFont(Path.of("/path/to/NotoSansJP-Regular.ttf"))
    .pdfPassword("user-password")           // required to open the PDF
    .pdfOwnerPassword("owner-password")     // optional; controls security settings
    .build();
```

- `pdfPassword` — the user password required to open the PDF.
- `pdfOwnerPassword` — the owner password that controls the PDF's security settings
  (e.g. adding print/copy restrictions with an external tool afterwards).
  When not set, `pdfPassword` is used as the owner password as well.
- `excelPassword` — separately, a password-protected **source** Excel file can be opened
  by setting this option.
