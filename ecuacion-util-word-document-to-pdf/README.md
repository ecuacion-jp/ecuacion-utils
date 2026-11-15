# ecuacion-util-word-document-to-pdf

## What is it?

`ecuacion-util-word-document-to-pdf` converts Word documents to PDF: `.docx`, and also `.docm`, `.dotx` and `.dotm` (macros are not run).

It is meant for text-centered documents such as contracts and terms of service, and outputs them as closely as possible to the appearance in Word:

- Paragraph formatting (alignment including justification, indents, spacing, line spacing, the document grid, tabs with leaders, borders, shading, keep-with-next, widow / orphan control)
- Character formatting (fonts per script, size, bold, italic, underline, strikethrough, color, highlight, superscript / subscript)
- Styles (document defaults, "based on" chains, theme fonts)
- Numbered headings and lists (multi-level, e.g. `第1条` / `1.1` / `(1)`)
- Tables (borders, shading, merged cells, repeated header rows, rows split across pages, nested tables)
- Sections (page size, margins), headers / footers, page numbers
- Tables of contents and cross-references: page numbers are **recalculated** from the layout of the PDF, and link to the headings
- Japanese line breaking rules (kinsoku, hanging punctuation, compression of punctuation, automatic space between Japanese and Latin text)
- Pictures (PNG / JPEG etc.), inline and floating (without text wrapping)
- Headings are also output as the outline (bookmarks) of the PDF
- Tracked changes (insertions are output, deletions are not), content controls, password-protected Word files

Not supported: shapes, text boxes, charts, SmartArt, equations, footnotes / endnotes, comments, multiple columns, vertical writing, text wrapping around floating pictures, EMF / WMF pictures, the legacy `.doc` format.

## Usage Example

```java
PdfGenerateOptions options =
    PdfGenerateOptions.builderForExplicitFont(Path.of("/path/to/NotoSansJP-Regular.ttf"))
    .addBoldFontPath(Path.of("/path/to/NotoSansJP-Bold.ttf"))
    .build();

WordToPdfUtil.generate(
    Path.of("contract.docx"),
    Path.of("contract.pdf"),
    options);
```

Use `PdfGenerateOptions.builderForSystemFonts()` instead to render with the fonts specified in the document, looked up from the fonts installed on the system.
The way fonts are resolved is the same as `ecuacion-util-excel-report-to-pdf`.

## Notes

- The layout is verified against PDF files exported by Word for Mac (line breaks, page breaks and line positions match). Word for Windows lays out some fonts (e.g. Yu Mincho of 11pt or larger on the document grid) with larger line heights, so page breaks can differ from a PDF exported by Word for Windows.
- To get the same layout as Word, install the fonts the documents use on the server and use `builderForSystemFonts()`. With other fonts the widths of the characters differ, and so do where lines and pages break.
- As in Word, a bold face is used only when the font family has one of its own (e.g. Yu Gothic, Arial, Meiryo). Otherwise (e.g. Yu Mincho, Century, MS Mincho) the regular face is emboldened, advancing 2% of the size more.
- Italic text is drawn by slanting the regular face.
- In justified paragraphs, punctuation and spaces between words may be compressed to put one more character or word on a line, as Word does.

## Exception Handling

`WordToPdfUtil.generate()` throws the following unchecked exceptions.

| Exception | Thrown when |
| --- | --- |
| `SystemFontNotFoundException` | System fonts are used, and the default font of the document is not installed with no fallback font registered. |
| `FontLoadFailedException` | A system font file was found but could not be loaded. |
| `CharacterNotRenderableException` | No configured font has the glyph of a character in the document. |
| `FileFormatNotSupportedException` | The file is not a Word file of a supported format (docx, docm, dotx, dotm), e.g. a legacy `.doc` file. |
| `FilePasswordRequiredException` | The Word file is password-protected, and `wordPassword` is not set. |
| `FilePasswordIncorrectException` | The Word file is password-protected, and `wordPassword` is wrong. |
| `java.io.UncheckedIOException` | A technical failure occurred while reading the Word file or writing the PDF file. |

All but the last extend `PdfGenerateException` (in `jp.ecuacion.util.pdf.common.exception`), and carry
a message ID for i18n display to the end user.

## Dependent Ecuacion Libraries

### Automatically Loaded Libraries

- `ecuacion-util-pdf-common`

### Manual Load Needed Libraries

- `ecuacion-lib-core`

## Dependent External Libraries

### Automatically Loaded Libraries

- `org.apache.pdfbox`
- `org.apache.poi:poi`
- `org.apache.poi:poi-ooxml`
- `jakarta.validation:jakarta.validation-api`
- `org.slf4j:slf4j-api`
- `org.apache.commons:commons-lang3`

### Manual Load Needed Libraries

- (If you use validation features, any `jakarta.validation:jakarta.validation-api` compatible Jakarta Validation implementation is needed. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- (any `org.slf4j:slf4j-api` compatible logging implementation. `ch.qos.logback:logback-classic` is recommended.)

## Documentation

- Official reference documentation - [ecuacion-references-utils](https://references.ecuacion.jp/ecuacion-references-utils/public/showMarkdown/page?id=word-document-to-pdf/overview)
- javadoc - [javadoc.io](https://javadoc.io/doc/jp.ecuacion.util/ecuacion-util-word-document-to-pdf/latest/jp.ecuacion.util.pdf.word.document/module-summary.html)

## Installation

Check [Installation](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
The description of dependent `ecuacion` modules is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-word-document-to-pdf</artifactId>
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
