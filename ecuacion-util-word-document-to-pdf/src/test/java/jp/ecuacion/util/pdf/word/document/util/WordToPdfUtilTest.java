/*
 * Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package jp.ecuacion.util.pdf.word.document.util;

import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.field;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.p;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.pageBreak;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.r;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.pdf.common.exception.CharacterNotRenderableException;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationLink;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("WordToPdfUtil")
class WordToPdfUtilTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  @SuppressWarnings("null")
  private static Path font(String name) throws URISyntaxException {
    return Path.of(WordToPdfUtilTest.class.getResource("/fonts/NotoSansJP/" + name).toURI());
  }

  private static PdfGenerateOptions options() throws URISyntaxException {
    return PdfGenerateOptions.builderForExplicitFont(font("NotoSansJP-Regular.ttf"))
        .addBoldFontPath(font("NotoSansJP-Bold.ttf")).build();
  }

  /** Converts the document and returns the text of each page. */
  private List<String> convert(DocxBuilder docx) throws Exception {
    Path pdf = generate(docx, options());
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      return pageTexts(doc);
    }
  }

  private Path generate(DocxBuilder docx, PdfGenerateOptions options) throws IOException {
    Path word = docx.write(tempDir.resolve("test.docx"));
    Path pdf = tempDir.resolve("test.pdf");
    WordToPdfUtil.generate(word, pdf, options);
    return pdf;
  }

  /**
   * Returns the text of each page without spaces: the text extraction reports the gaps of
   * the layout (e.g. the automatic space between Japanese text and digits) as spaces.
   */
  private static List<String> pageTexts(PDDocument doc) throws IOException {
    List<String> pages = new ArrayList<>();
    for (int i = 1; i <= doc.getNumberOfPages(); i++) {
      PDFTextStripper stripper = new PDFTextStripper();
      stripper.setStartPage(i);
      stripper.setEndPage(i);
      pages.add(stripper.getText(doc).replace(" ", ""));
    }
    return pages;
  }

  private static String numberingOfArticles() {
    return "<w:abstractNum w:abstractNumId=\"0\">"
        + "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>"
        + "<w:suff w:val=\"space\"/><w:lvlText w:val=\"第%1条\"/></w:lvl>"
        + "<w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/>"
        + "<w:lvlText w:val=\"(%2)\"/><w:pPr><w:ind w:left=\"420\" w:hanging=\"420\"/></w:pPr>"
        + "</w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
  }

  private static String numbered(int level, String text) {
    return p("<w:numPr><w:ilvl w:val=\"" + level + "\"/><w:numId w:val=\"1\"/></w:numPr>",
        r(text));
  }

  @Nested
  @DisplayName("paragraphs and pages")
  class ParagraphsAndPages {

    @Test
    @DisplayName("paragraphs are output in order on one page")
    void paragraphsInOrder() throws Exception {
      List<String> pages = convert(new DocxBuilder().add(p(r("最初の段落です。")))
          .add(p(r("Second paragraph."))).add(p(r("三番目。"))));

      assertThat(pages).hasSize(1);
      assertThat(pages.get(0)).containsSubsequence("最初の段落です。", "Secondparagraph.", "三番目。");
    }

    @Test
    @DisplayName("an empty document produces one empty page")
    void emptyDocument() throws Exception {
      List<String> pages = convert(new DocxBuilder());

      assertThat(pages).hasSize(1);
      assertThat(pages.get(0)).isBlank();
    }

    @Test
    @DisplayName("a page break run starts a new page")
    void explicitPageBreak() throws Exception {
      List<String> pages = convert(
          new DocxBuilder().add(p(r("一ページ目") + pageBreak() + r("二ページ目"))));

      assertThat(pages).hasSize(2);
      assertThat(pages.get(0)).contains("一ページ目").doesNotContain("二ページ目");
      assertThat(pages.get(1)).contains("二ページ目");
    }

    @Test
    @DisplayName("pageBreakBefore starts a new page")
    void pageBreakBefore() throws Exception {
      List<String> pages = convert(new DocxBuilder().add(p(r("前")))
          .add(p("<w:pageBreakBefore/>", r("後"))));

      assertThat(pages).hasSize(2);
      assertThat(pages.get(1)).contains("後");
    }

    @Test
    @DisplayName("content which does not fit on a page flows to the next pages")
    void flowsToNextPages() throws Exception {
      DocxBuilder docx = new DocxBuilder();
      for (int i = 1; i <= 120; i++) {
        docx.add(p(r("段落" + i + "番")));
      }
      List<String> pages = convert(docx);

      assertThat(pages.size()).isGreaterThan(1);
      assertThat(pages.get(0)).contains("段落1番");
      assertThat(pages.get(pages.size() - 1)).contains("段落120番");
    }

    @SuppressWarnings("null")
    @Test
    @DisplayName("a long paragraph is wrapped, and no line starts with closing punctuation")
    void wrappingAndKinsoku() throws Exception {
      String sentence = "これは禁則処理を確認するための文です。";
      List<String> pages = convert(new DocxBuilder().add(p(r(sentence.repeat(30)))));

      List<String> lines = pages.get(0).lines().filter(l -> !l.isBlank()).toList();
      assertThat(lines.size()).isGreaterThan(5);
      assertThat(lines).allSatisfy(l -> assertThat(l).doesNotStartWith("。"));
    }

    @Test
    @DisplayName("a section with another page size produces pages of that size")
    void sections() throws Exception {
      String landscape = "<w:sectPr><w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>"
          + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>"
          + "</w:sectPr>";
      Path pdf = generate(new DocxBuilder().add(p(landscape, r("横"))).add(p(r("縦"))),
          options());

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        assertThat(doc.getPage(0).getMediaBox().getWidth())
            .isGreaterThan(doc.getPage(0).getMediaBox().getHeight());
        assertThat(doc.getPage(1).getMediaBox().getWidth())
            .isLessThan(doc.getPage(1).getMediaBox().getHeight());
      }
    }
  }

  @Nested
  @DisplayName("character formatting")
  class CharacterFormatting {

    @Test
    @DisplayName("hidden text and deleted text (tracked changes) are not output")
    void hiddenAndDeleted() throws Exception {
      List<String> pages = convert(new DocxBuilder().add(p(r("表示") + r("非表示", "<w:vanish/>")
          + "<w:del w:id=\"1\"><w:r><w:delText>削除済</w:delText></w:r></w:del>"
          + "<w:ins w:id=\"2\">" + r("挿入済") + "</w:ins>")));

      assertThat(pages.get(0)).contains("表示", "挿入済").doesNotContain("非表示", "削除済");
    }

    @Test
    @DisplayName("bold, italic, underline, superscript and all caps are output")
    void decorations() throws Exception {
      List<String> pages = convert(new DocxBuilder().add(p(r("太字", "<w:b/>")
          + r("斜体", "<w:i/>") + r("下線", "<w:u w:val=\"single\"/>")
          + r("2", "<w:vertAlign w:val=\"superscript\"/>") + r("caps", "<w:caps/>"))));

      assertThat(pages.get(0)).contains("太字", "斜体", "下線", "CAPS");
    }

    @Test
    @DisplayName("a character no font has: CharacterNotRenderableException")
    void characterNotRenderable() {
      DocxBuilder docx = new DocxBuilder().add(p(r("😀")));

      assertThatThrownBy(() -> generate(docx, options()))
          .isInstanceOf(CharacterNotRenderableException.class);
    }
  }

  @Nested
  @DisplayName("numbering")
  class NumberingTest {

    @Test
    @DisplayName("multi-level numbers are counted, and a lower level restarts")
    void multiLevel() throws Exception {
      List<String> pages = convert(new DocxBuilder().numbering(numberingOfArticles())
          .add(numbered(0, "目的")).add(numbered(1, "一号")).add(numbered(1, "二号"))
          .add(numbered(0, "定義")).add(numbered(1, "再開")));

      assertThat(pages.get(0)).containsSubsequence("第1条", "目的", "(1)", "一号", "(2)", "二号",
          "第2条", "定義", "(1)", "再開");
    }

    @Test
    @DisplayName("the numbering of a paragraph style is applied")
    void numberingFromStyle() throws Exception {
      String styles = "<w:style w:type=\"paragraph\" w:styleId=\"H1\"><w:name w:val=\"heading 1\"/>"
          + "<w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>";
      List<String> pages = convert(new DocxBuilder().numbering(numberingOfArticles())
          .styles(styles).add(p("<w:pStyle w:val=\"H1\"/>", r("総則")))
          .add(p("<w:pStyle w:val=\"H1\"/>", r("雑則"))));

      assertThat(pages.get(0)).containsSubsequence("第1条", "総則", "第2条", "雑則");
    }
  }

  @Nested
  @DisplayName("fields")
  class Fields {

    @Test
    @DisplayName("PAGE and NUMPAGES in a footer show the page number and the page count")
    void pageNumbers() throws Exception {
      List<String> pages = convert(new DocxBuilder()
          .footer(p(field("PAGE", "9") + r(" / ") + field("NUMPAGES", "9")))
          .add(p(r("A") + pageBreak() + r("B") + pageBreak() + r("C"))));

      assertThat(pages).hasSize(3);
      assertThat(pages.get(0)).contains("1/3");
      assertThat(pages.get(1)).contains("2/3");
      assertThat(pages.get(2)).contains("3/3");
    }

    @Test
    @DisplayName("a header is output on every page")
    void header() throws Exception {
      List<String> pages = convert(new DocxBuilder().header(p(r("契約書ヘッダー")))
          .add(p(r("A") + pageBreak() + r("B"))));

      assertThat(pages).allSatisfy(t -> assertThat(t).contains("契約書ヘッダー"));
    }

    @Test
    @DisplayName("page numbers in a table of contents are recalculated, and link to headings")
    void tableOfContents() throws Exception {
      String toc = p("<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9000\"/></w:tabs>",
          "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC \\o </w:instrText>"
              + "</w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
              + "<w:hyperlink w:anchor=\"_Toc1\">" + r("見出し") + "<w:r><w:tab/></w:r>"
              + field("PAGEREF _Toc1 \\h", "99") + "</w:hyperlink>"
              + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
      Path pdf = generate(new DocxBuilder().add(toc).add(p(pageBreak())).add(p(pageBreak()))
          .add(p("<w:outlineLvl w:val=\"0\"/>",
              "<w:bookmarkStart w:id=\"1\" w:name=\"_Toc1\"/>" + r("見出し本体")
                  + "<w:bookmarkEnd w:id=\"1\"/>")),
          options());

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        List<String> pages = pageTexts(doc);
        assertThat(pages).hasSize(3);
        assertThat(pages.get(0)).contains("見出し").contains("3").doesNotContain("99");
        assertThat(pages.get(2)).contains("見出し本体");
        assertThat(doc.getPage(0).getAnnotations()).anyMatch(a -> a instanceof PDAnnotationLink);
        PDOutlineItem item = doc.getDocumentCatalog().getDocumentOutline().getFirstChild();
        assertThat(item.getTitle()).isEqualTo("見出し本体");
        assertThat(doc.getPages().indexOf(item.findDestinationPage(doc))).isEqualTo(2);
      }
    }
  }

  @Nested
  @DisplayName("tables")
  class Tables {

    private static String cell(String text) {
      return "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>" + p(r(text))
          + "</w:tc>";
    }

    private static String table(String rows) {
      return "<w:tbl><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\"/>"
          + "<w:insideH w:val=\"single\" w:sz=\"4\"/></w:tblBorders></w:tblPr><w:tblGrid>"
          + "<w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>" + rows + "</w:tbl>";
    }

    @Test
    @DisplayName("the text of cells is output row by row")
    void cells() throws Exception {
      List<String> pages = convert(new DocxBuilder().add(table(
          "<w:tr>" + cell("区分") + cell("金額") + "</w:tr><w:tr>" + cell("基本") + cell("100円")
              + "</w:tr>")));

      assertThat(pages.get(0)).containsSubsequence("区分", "金額", "基本", "100円");
    }

    @Test
    @DisplayName("header rows are repeated on every page the table continues to")
    void repeatedHeaderRows() throws Exception {
      StringBuilder rows = new StringBuilder(
          "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + cell("見出し行") + cell("値") + "</w:tr>");
      for (int i = 1; i <= 100; i++) {
        rows.append("<w:tr>").append(cell("行" + i)).append(cell("v")).append("</w:tr>");
      }
      List<String> pages = convert(new DocxBuilder().add(table(rows.toString())));

      assertThat(pages.size()).isGreaterThan(1);
      assertThat(pages).allSatisfy(t -> assertThat(t).contains("見出し行"));
      assertThat(pages.get(pages.size() - 1)).contains("行100");
    }

    @Test
    @DisplayName("a row taller than the rest of the page is split across pages")
    void rowSplit() throws Exception {
      StringBuilder paragraphs = new StringBuilder();
      for (int i = 1; i <= 80; i++) {
        paragraphs.append(p(r("セル内" + i + "行")));
      }
      String row = "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"6000\" w:type=\"dxa\"/></w:tcPr>"
          + paragraphs + "</w:tc></w:tr>";
      List<String> pages = convert(new DocxBuilder().add(p(r("前文"))).add(table(row)));

      assertThat(pages.size()).isGreaterThan(1);
      assertThat(pages.get(0)).contains("前文", "セル内1行");
      assertThat(pages.get(pages.size() - 1)).contains("セル内80行");
    }
  }

  @Nested
  @DisplayName("files and passwords")
  class FilesAndPasswords {

    @Test
    @DisplayName("pdfPassword: the PDF file is encrypted, and opens with the password")
    void pdfPassword() throws Exception {
      PdfGenerateOptions options =
          PdfGenerateOptions.builderForExplicitFont(font("NotoSansJP-Regular.ttf"))
              .pdfPassword("secret").build();
      Path pdf = generate(new DocxBuilder().add(p(r("秘密"))), options);

      assertThatThrownBy(() -> Loader.loadPDF(pdf.toFile())).isInstanceOf(IOException.class);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile(), "secret")) {
        assertThat(doc.isEncrypted()).isTrue();
        assertThat(pageTexts(doc).get(0)).contains("秘密");
      }
    }

    @Test
    @DisplayName("a missing Word file: UncheckedIOException")
    void missingFile() {
      assertThatThrownBy(() -> WordToPdfUtil.generate(tempDir.resolve("none.docx"),
          tempDir.resolve("out.pdf"), options())).isInstanceOf(UncheckedIOException.class);
    }

    @Test
    @DisplayName("a file which is not a .docx file: an exception, and no PDF file is left")
    void notDocx() throws Exception {
      Path text = Files.writeString(tempDir.resolve("plain.docx"), "not a word file");
      Path pdf = tempDir.resolve("out.pdf");

      assertThatThrownBy(() -> WordToPdfUtil.generate(text, pdf, options()))
          .isInstanceOf(RuntimeException.class);
      assertThat(pdf).doesNotExist();
    }
  }
}
