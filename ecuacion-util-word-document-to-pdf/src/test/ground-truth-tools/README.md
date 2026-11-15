# Ground truth tools

The files in `src/test/resources/ground-truth` are pairs of a Word file (`<name>.docx`) and the
PDF file Microsoft Word for Mac exported from it (`<name>.docx.pdf`). `WordGroundTruthTest`
converts every Word file in the directory with the fonts of the system, and checks that every line
has the same text at the same position (within 1pt) as in Word's PDF. The test is skipped where
the fonts the files use (Yu Mincho etc.) are not installed, e.g. on CI.

When changing or adding layout behavior, do not implement it from a guess: make a Word file which
has the case, export it with Word, and compare. Most of the differences found so far could not be
seen by eye.

## Files

| File | Content |
| --- | --- |
| `docx.py` | Common builder: makes a minimal `.docx` file Word accepts from WordprocessingML fragments (`build()`, `p()`, `r()`, `fld()`, `fonts()`, `default_styles()`). Supports headers / footers of several sections (`stories`) and PNG pictures (`media`) |
| `export.sh` | Exports a `.docx` file to PDF with Word for Mac, driven by AppleScript |
| `gen_contract.py` | A sample contract: title, table of contents, numbered articles, a table, signature lines (args: `<out.docx> [East Asian font]`) |
| `gen_lines.py` | Line heights of fonts, sizes and line spacing rules (args: `<out.docx> grid` or `<out.docx> nogrid`) |
| `gen_paragraphs.py` | Indents, alignment, tabs, numbering formats, paragraph spacing, borders, East Asian spacing |
| `gen_pagination.py` | Keep with next, widow / orphan control, keep lines, space before at page tops, tables across pages |
| `gen_tables.py` | Table widths, position, margins, borders, merged cells, row heights, table styles, nested tables |
| `gen_english.py` | Justified English text (how far Word compresses spaces to put one more word on a line) |
| `gen_characters.py` | Character formatting, fonts per script, bold faces, precedence of styles |
| `gen_sections.py` | Sections, page settings, first / even page headers and footers, page number formats |
| `gen_images.py` | Inline pictures and the heights of their lines |

Every generator but `gen_lines.py` takes only the output path.

## How to add or update a pair (on a Mac with Microsoft Word)

```bash
cd ecuacion-util-word-document-to-pdf/src/test/ground-truth-tools
python3 gen_tables.py /tmp/tables.docx
bash export.sh /tmp/tables.docx /tmp/tables.docx.pdf
cp /tmp/tables.docx /tmp/tables.docx.pdf ../resources/ground-truth/
```

(A file written by Word itself can be used as well.) The test picks up every `.docx` file in the
directory, so nothing has to be added to the test.

## Notes

- Word is sandboxed, so `export.sh` passes the files through the Office group container
  (`~/Library/Group Containers/UBF8T346G9.Office/ecuacion-word-gt`), which Word can read and
  write without asking for file access.
- Word rejects a file whose elements are not in the order of the schema (e.g. `w:spacing` must
  come before `w:ind` in `w:pPr`, `w:color` before `w:sz` in `w:rPr`) as "possibly corrupted".
- The comparison normalizes the text: Word's PDF files use Kangxi radicals and compatibility
  characters for some kanji (e.g. `⽬` for `目`), and the number of the characters of tab leaders
  is not compared.
