# ecuacion-util-pdf-common

## What is it?

`ecuacion-util-pdf-common` provides the features shared by the `ecuacion-util` modules which generate PDF files
(`ecuacion-util-excel-report-to-pdf` and `ecuacion-util-word-document-to-pdf`): font resolution
(system font lookup, fallback fonts) and the exceptions related to it.

It is not meant to be used alone. It is loaded automatically as a dependency of the modules above,
and the only part of it which applications use directly is the exception classes in
`jp.ecuacion.util.pdf.common.exception`.

| Exception | Thrown when |
| --- | --- |
| `PdfGenerateException` | (The abstract superclass of the exceptions thrown by PDF generation.) |
| `SystemFontNotFoundException` | System fonts are used, and the default font of the source file is not installed with no fallback font registered. |
| `FontLoadFailedException` | A system font file was found but could not be loaded. |
| `CharacterNotRenderableException` | No configured font has the glyph of a character in the source file. |

## Documentation

- Official reference documentation - [ecuacion-references-utils](https://references.ecuacion.jp/ecuacion-references-utils/public/showMarkdown/page?id=home)
- javadoc - [javadoc.io](https://javadoc.io/doc/jp.ecuacion.util/ecuacion-util-pdf-common/latest/jp.ecuacion.util.pdf.common/module-summary.html)
