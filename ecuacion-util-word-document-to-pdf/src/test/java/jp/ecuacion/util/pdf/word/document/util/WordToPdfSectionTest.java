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
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Tests page settings, sections, headers / footers and page numbers. */
@DisplayName("WordToPdfUtil: sections, headers and footers")
class WordToPdfSectionTest extends AbstractWordToPdfTest {

  private static final String A4 = "<w:pgSz w:w=\"11906\" w:h=\"16838\"/>";
  private static final String MARGINS = "<w:pgMar w:top=\"1440\" w:right=\"1440\" "
      + "w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/>";

  private static String headerRef(String type, String relId) {
    return "<w:headerReference w:type=\"" + type + "\" r:id=\"" + relId + "\"/>";
  }

  private static String footerRef(String type, String relId) {
    return "<w:footerReference w:type=\"" + type + "\" r:id=\"" + relId + "\"/>";
  }

  /** Returns a paragraph which ends a section with the section properties. */
  private static String sectionEnd(String text, String sectPr) {
    return p("<w:sectPr>" + sectPr + "</w:sectPr>", r(text));
  }

  /** Returns the texts of the lines of each page, joined with "|". */
  private static String text(List<Line> page) {
    StringBuilder sb = new StringBuilder();
    for (Line line : page) {
      sb.append(line.text()).append('|');
    }
    return sb.toString();
  }

  @Nested
  @DisplayName("page settings")
  class PageSettings {

    @Test
    @DisplayName("page size and orientation")
    void pageSize() throws Exception {
      Path pdf = generate(new DocxBuilder()
          .sectPr("<w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>" + MARGINS)
          .add(p(r("横"))));

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getPage(0).getMediaBox().getWidth()).isCloseTo(841.9f, TOL);
        assertThat(doc.getPage(0).getMediaBox().getHeight()).isCloseTo(595.3f, TOL);
      }
    }

    @Test
    @DisplayName("margins move the text area")
    void margins() throws Exception {
      Line normal = convert(new DocxBuilder().add(p("<w:jc w:val=\"both\"/>", r(chars(80)))))
          .get(0).get(0);
      Line wide = convert(new DocxBuilder().sectPr(A4 + "<w:pgMar w:top=\"2880\" "
          + "w:right=\"2160\" w:bottom=\"1440\" w:left=\"2880\" w:header=\"720\" "
          + "w:footer=\"720\"/>").add(p("<w:jc w:val=\"both\"/>", r(chars(80))))).get(0).get(0);

      assertThat(wide.left()).isCloseTo(144f, TOL);
      assertThat(wide.right()).isCloseTo(595.3f - 108, TOL);
      assertThat(wide.baseline() - normal.baseline()).isCloseTo(72f, TOL);
    }

    @Test
    @DisplayName("the gutter is added to the left margin")
    void gutter() throws Exception {
      Line line = convert(new DocxBuilder().sectPr(A4 + "<w:pgMar w:top=\"1440\" "
          + "w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:gutter=\"720\"/>")
          .add(p(r("あ")))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT + 36, TOL);
    }

    @Test
    @DisplayName("the bottom margin decides how many lines fit on a page")
    void bottomMargin() throws Exception {
      DocxBuilder normal = new DocxBuilder();
      DocxBuilder large = new DocxBuilder().sectPr(A4 + "<w:pgMar w:top=\"1440\" "
          + "w:right=\"1440\" w:bottom=\"7200\" w:left=\"1440\"/>");
      for (int i = 0; i < 60; i++) {
        normal.add(p(r("行")));
        large.add(p(r("行")));
      }

      assertThat(convert(large).get(0).size()).isLessThan(convert(normal).get(0).size() - 10);
    }
  }

  @Nested
  @DisplayName("sections")
  class Sections {

    @Test
    @DisplayName("a section starts on a new page, with its own page size")
    void nextPage() throws Exception {
      Path pdf = generate(new DocxBuilder().add(sectionEnd("一", A4 + MARGINS))
          .add(p(r("二")))
          .sectPr("<w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>" + MARGINS));

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        assertThat(doc.getPage(0).getMediaBox().getWidth()).isCloseTo(595.3f, TOL);
        assertThat(doc.getPage(1).getMediaBox().getWidth()).isCloseTo(841.9f, TOL);
      }
    }

    @Test
    @DisplayName("a continuous section continues on the same page")
    void continuous() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(sectionEnd("一", A4 + MARGINS))
          .add(p(r("二"))).sectPr(A4 + MARGINS + "<w:type w:val=\"continuous\"/>"));

      assertThat(pages).hasSize(1);
      assertThat(text(pages.get(0))).isEqualTo("一|二|");
    }

    @Test
    @DisplayName("a continuous section with another page size starts on a new page")
    void continuousWithAnotherPageSize() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(sectionEnd("一", A4 + MARGINS))
          .add(p(r("二"))).sectPr("<w:pgSz w:w=\"16838\" w:h=\"11906\"/>" + MARGINS
              + "<w:type w:val=\"continuous\"/>"));

      assertThat(pages).hasSize(2);
    }

    @Test
    @DisplayName("each section has its own document grid")
    void gridPerSection() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder()
          .add(p(r("一" + chars(100))))
          .add(sectionEnd("区切り", A4 + MARGINS
              + "<w:docGrid w:type=\"lines\" w:linePitch=\"600\"/>"))
          .add(p(r("二" + chars(100)))).sectPr(A4 + MARGINS));
      List<Line> first = pages.get(0);
      List<Line> second = pages.get(1);

      assertThat(first.get(1).baseline() - first.get(0).baseline()).isCloseTo(30f, TOL);
      assertThat(second.get(1).baseline() - second.get(0).baseline()).isCloseTo(basePitch(),
          TOL);
    }
  }

  @Nested
  @DisplayName("headers and footers")
  class HeadersAndFooters {

    private DocxBuilder threePages() {
      return new DocxBuilder().add(p(r("本文一") + pageBreak() + r("本文二") + pageBreak()
          + r("本文三")));
    }

    @Test
    @DisplayName("the header is at the header distance, the footer above the footer distance")
    void positions() throws Exception {
      final float base = basePitch();
      List<Line> page = convert(new DocxBuilder().header(p(r("ヘッダー")))
          .footer(p(r("フッター一")) + p(r("フッター二"))).add(p(r("本文")))).get(0);
      List<Line> plain = convert(new DocxBuilder().add(p(r("本文")))).get(0);
      float firstLineOffset = plain.get(0).baseline() - 72;

      assertThat(page.get(0).text()).isEqualTo("ヘッダー");
      assertThat(page.get(0).baseline()).isCloseTo(36 + firstLineOffset, TOL);
      assertThat(page.get(3).text()).isEqualTo("フッター二");
      assertThat(page.get(3).baseline()).isCloseTo(841.9f - 36 - base + firstLineOffset, TOL);
      assertThat(page.get(3).baseline() - page.get(2).baseline()).isCloseTo(base, TOL);
    }

    @Test
    @DisplayName("a header taller than the top margin pushes the body down")
    void tallHeader() throws Exception {
      final float base = basePitch();
      StringBuilder header = new StringBuilder();
      for (int i = 0; i < 6; i++) {
        header.append(p(r("ヘッダー")));
      }
      List<List<Line>> pages = convert(new DocxBuilder().header(header.toString())
          .add(p(r("本文"))));
      float normal = convert(new DocxBuilder().add(p(r("本文")))).get(0).get(0).baseline();

      // The header ends at 36 + 6 lines, below the top margin (72).
      assertThat(PdfLines.find(pages, "本文").baseline() - normal)
          .isCloseTo(36 + base * 6 - 72, TOL);
    }

    @Test
    @DisplayName("a footer taller than the bottom margin makes the body shorter")
    void tallFooter() throws Exception {
      StringBuilder footer = new StringBuilder();
      for (int i = 0; i < 10; i++) {
        footer.append(p(r("フッター")));
      }
      DocxBuilder normal = new DocxBuilder();
      DocxBuilder withFooter = new DocxBuilder().footer(footer.toString());
      for (int i = 0; i < 60; i++) {
        normal.add(p(r("行")));
        withFooter.add(p(r("行")));
      }
      long bodyLines = convert(withFooter).get(0).stream()
          .filter(l -> l.text().equals("行")).count();

      assertThat(bodyLines).isLessThan(convert(normal).get(0).size() - 5);
    }

    @Test
    @DisplayName("titlePg: the first page uses the first-page header")
    void firstPage() throws Exception {
      List<List<Line>> pages = convert(threePages().story("rH1", true, p(r("既定")))
          .story("rH2", true, p(r("先頭")))
          .sectPr(headerRef("default", "rH1") + headerRef("first", "rH2") + A4 + MARGINS
              + "<w:titlePg/>"));

      assertThat(text(pages.get(0))).isEqualTo("先頭|本文一|");
      assertThat(text(pages.get(1))).isEqualTo("既定|本文二|");
      assertThat(text(pages.get(2))).isEqualTo("既定|本文三|");
    }

    @Test
    @DisplayName("titlePg without a first-page header: the first page has no header")
    void firstPageWithoutHeader() throws Exception {
      List<List<Line>> pages = convert(threePages().story("rH1", true, p(r("既定")))
          .sectPr(headerRef("default", "rH1") + A4 + MARGINS + "<w:titlePg/>"));

      assertThat(text(pages.get(0))).isEqualTo("本文一|");
      assertThat(text(pages.get(1))).isEqualTo("既定|本文二|");
    }

    @Test
    @DisplayName("without titlePg the first-page header is not used")
    void firstPageHeaderIgnored() throws Exception {
      List<List<Line>> pages = convert(threePages().story("rH1", true, p(r("既定")))
          .story("rH2", true, p(r("先頭")))
          .sectPr(headerRef("default", "rH1") + headerRef("first", "rH2") + A4 + MARGINS));

      assertThat(text(pages.get(0))).isEqualTo("既定|本文一|");
    }

    @Test
    @DisplayName("evenAndOddHeaders: even pages use the even header and footer")
    void evenAndOdd() throws Exception {
      List<List<Line>> pages = convert(threePages().settings("<w:evenAndOddHeaders/>")
          .story("rH1", true, p(r("奇数"))).story("rH2", true, p(r("偶数")))
          .story("rF2", false, p(r("偶数フッター")))
          .sectPr(headerRef("default", "rH1") + headerRef("even", "rH2")
              + footerRef("even", "rF2") + A4 + MARGINS));

      assertThat(text(pages.get(0))).isEqualTo("奇数|本文一|");
      assertThat(text(pages.get(1))).isEqualTo("偶数|本文二|偶数フッター|");
      assertThat(text(pages.get(2))).isEqualTo("奇数|本文三|");
    }

    @Test
    @DisplayName("without evenAndOddHeaders the even header is not used")
    void evenIgnored() throws Exception {
      List<List<Line>> pages = convert(threePages().story("rH1", true, p(r("奇数")))
          .story("rH2", true, p(r("偶数")))
          .sectPr(headerRef("default", "rH1") + headerRef("even", "rH2") + A4 + MARGINS));

      assertThat(text(pages.get(1))).isEqualTo("奇数|本文二|");
    }

    @Test
    @DisplayName("a section without a header inherits it, a section with one replaces it")
    void inheritance() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().story("rH1", true, p(r("一の頭")))
          .story("rH3", true, p(r("三の頭")))
          .add(sectionEnd("一", headerRef("default", "rH1") + A4 + MARGINS))
          .add(sectionEnd("二", A4 + MARGINS)).add(p(r("三")))
          .sectPr(headerRef("default", "rH3") + A4 + MARGINS));

      assertThat(text(pages.get(0))).isEqualTo("一の頭|一|");
      assertThat(text(pages.get(1))).isEqualTo("一の頭|二|");
      assertThat(text(pages.get(2))).isEqualTo("三の頭|三|");
    }

    @Test
    @DisplayName("a table in a header is output")
    void tableInHeader() throws Exception {
      String table = "<w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/>"
          + "</w:tblGrid><w:tr><w:tc>" + p(r("左")) + "</w:tc><w:tc>" + p(r("右"))
          + "</w:tc></w:tr></w:tbl>";
      List<List<Line>> pages = convert(new DocxBuilder().header(table + p(""))
          .add(p(r("本文"))));

      assertThat(pages.get(0).get(0).text()).isEqualTo("左右");
    }
  }

  @Nested
  @DisplayName("page numbers")
  class PageNumbers {

    private DocxBuilder numbered(String pgNumType) {
      return new DocxBuilder().add(p(r("一") + pageBreak() + r("二") + pageBreak() + r("三")))
          .story("rF1", false, p(r("[") + field("PAGE", "9") + r("]")))
          .sectPr(footerRef("default", "rF1") + A4 + MARGINS + pgNumType);
    }

    @ParameterizedTest(name = "{0}")
    @CsvSource({"decimal, [1], [2], [3]", "upperRoman, [I], [II], [III]",
        "lowerRoman, [i], [ii], [iii]", "upperLetter, [A], [B], [C]",
        "lowerLetter, [a], [b], [c]", "decimalFullWidth, [１], [２], [３]",
        "decimalEnclosedCircle, [①], [②], [③]"})
    @DisplayName("number format of the section")
    void format(String format, String first, String second, String third) throws Exception {
      List<List<Line>> pages = convert(numbered("<w:pgNumType w:fmt=\"" + format + "\"/>"));

      assertThat(text(pages.get(0))).endsWith(first + "|");
      assertThat(text(pages.get(1))).endsWith(second + "|");
      assertThat(text(pages.get(2))).endsWith(third + "|");
    }

    @Test
    @DisplayName("start number of the section")
    void start() throws Exception {
      List<List<Line>> pages = convert(numbered("<w:pgNumType w:start=\"5\"/>"));

      assertThat(text(pages.get(0))).endsWith("[5]|");
      assertThat(text(pages.get(2))).endsWith("[7]|");
    }

    @Test
    @DisplayName("the next section continues the numbers, or restarts with a start number")
    void acrossSections() throws Exception {
      String footer = footerRef("default", "rF1");
      DocxBuilder continued = new DocxBuilder().story("rF1", false, p(field("PAGE", "9")))
          .add(sectionEnd("一", footer + A4 + MARGINS)).add(p(r("二") + pageBreak() + r("三")))
          .sectPr(A4 + MARGINS);
      DocxBuilder restarted = new DocxBuilder().story("rF1", false, p(field("PAGE", "9")))
          .add(sectionEnd("一", footer + A4 + MARGINS)).add(p(r("二") + pageBreak() + r("三")))
          .sectPr(A4 + MARGINS + "<w:pgNumType w:start=\"1\"/>");

      assertThat(text(convert(continued).get(2))).isEqualTo("三|3|");
      assertThat(text(convert(restarted).get(1))).isEqualTo("二|1|");
      assertThat(text(convert(restarted).get(2))).isEqualTo("三|2|");
    }

    @Test
    @DisplayName("a format switch of the PAGE field overrides the format of the section")
    void fieldSwitch() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("一")))
          .footer(p(field("PAGE \\* ROMAN", "9") + r("/") + field("PAGE \\* roman", "9"))));

      assertThat(text(pages.get(0))).isEqualTo("一|I/i|");
    }

    @Test
    @DisplayName("SECTIONPAGES is the number of the pages of the section")
    void sectionPages() throws Exception {
      String footer = footerRef("default", "rF1");
      List<List<Line>> pages = convert(new DocxBuilder()
          .story("rF1", false, p(field("SECTIONPAGES", "9") + r("/") + field("NUMPAGES", "9")))
          .add(sectionEnd("一", footer + A4 + MARGINS))
          .add(p(r("二") + pageBreak() + r("三"))).sectPr(A4 + MARGINS));

      assertThat(text(pages.get(0))).isEqualTo("一|1/3|");
      assertThat(text(pages.get(1))).isEqualTo("二|2/3|");
      assertThat(text(pages.get(2))).isEqualTo("三|2/3|");
    }

    @Test
    @DisplayName("NUMPAGES and PAGE in the body")
    void fieldsInBody() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("一") + pageBreak()
          + r("第") + field("PAGE", "9") + r("頁、全") + field("NUMPAGES", "9") + r("頁"))));

      assertThat(text(pages.get(1))).isEqualTo("第2頁、全2頁|");
    }

    @Test
    @DisplayName("a simple field (fldSimple) works as well")
    void simpleField() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("一")))
          .footer(p("<w:fldSimple w:instr=\" PAGE \">" + r("9") + "</w:fldSimple>")));

      assertThat(text(pages.get(0))).isEqualTo("一|1|");
    }

    @Test
    @DisplayName("a field which is not calculated here shows the result stored in the file")
    void otherField() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder()
          .add(p(r("日付：") + field("DATE \\@ \"yyyy/MM/dd\"", "2026/04/01"))));

      assertThat(text(pages.get(0))).isEqualTo("日付：2026/04/01|");
    }
  }
}
