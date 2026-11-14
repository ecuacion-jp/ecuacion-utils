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

import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.p;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.r;
import static org.assertj.core.api.Assertions.assertThat;
import java.util.List;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

/**
 * Tests where pages break. Each test fills a page with one-line paragraphs, leaving some
 * lines at the bottom, and puts the paragraphs under test there.
 */
@DisplayName("WordToPdfUtil: page breaks")
class WordToPdfPaginationTest extends AbstractWordToPdfTest {

  private static final String NO_WIDOW = "<w:widowControl w:val=\"0\"/>";

  /** Returns a document whose first page has the number of lines left at the bottom. */
  private DocxBuilder leaving(int lines) throws Exception {
    int perPage = (int) Math.floor((BODY_HEIGHT + 0.01f) / basePitch());
    DocxBuilder docx = new DocxBuilder();
    for (int i = 1; i <= perPage - lines; i++) {
      docx.add(p(r("行" + i + "。")));
    }
    return docx;
  }

  @Nested
  @DisplayName("keep with next")
  class KeepWithNext {

    @Test
    @DisplayName("without it, a heading can be the last line of a page")
    void notSet() throws Exception {
      List<List<Line>> pages = convert(leaving(1).add(p(r("見出し"))).add(p(r("本文"))));

      assertThat(pageOf(pages, "見出し")).isEqualTo(0);
      assertThat(pageOf(pages, "本文")).isEqualTo(1);
    }

    @Test
    @DisplayName("a heading moves to the next page with the paragraph after it")
    void heading() throws Exception {
      List<List<Line>> pages =
          convert(leaving(1).add(p("<w:keepNext/>", r("見出し"))).add(p(r("本文"))));

      assertThat(pageOf(pages, "見出し")).isEqualTo(1);
      assertThat(pageOf(pages, "本文")).isEqualTo(1);
    }

    @Test
    @DisplayName("consecutive headings move together")
    void chain() throws Exception {
      List<List<Line>> pages = convert(leaving(2).add(p("<w:keepNext/>", r("見出し一")))
          .add(p("<w:keepNext/>", r("見出し二"))).add(p(r("本文"))));

      assertThat(pageOf(pages, "見出し一")).isEqualTo(1);
      assertThat(pageOf(pages, "見出し二")).isEqualTo(1);
    }

    @Test
    @DisplayName("the heading stays when the first lines of the next paragraph fit")
    void fits() throws Exception {
      List<List<Line>> pages = convert(
          leaving(2).add(p("<w:keepNext/>", r("見出し"))).add(p(NO_WIDOW, r("本" + chars(60)))));

      assertThat(pageOf(pages, "見出し")).isEqualTo(0);
      assertThat(pageOf(pages, "本")).isEqualTo(0);
    }

    @Test
    @DisplayName("a heading moves to the next page with the table after it")
    void beforeTable() throws Exception {
      String table = "<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>"
          + p(r("表")) + "</w:tc></w:tr></w:tbl>";
      List<List<Line>> pages =
          convert(leaving(1).add(p("<w:keepNext/>", r("見出し"))).add(table));

      assertThat(pageOf(pages, "見出し")).isEqualTo(1);
    }
  }

  @Nested
  @DisplayName("widow / orphan control and keep lines")
  class WidowControl {

    @Test
    @DisplayName("off: a paragraph is divided wherever the page ends")
    void off() throws Exception {
      List<List<Line>> pages = convert(leaving(1).add(p(NO_WIDOW, r("頭" + chars(99)))));

      assertThat(pageOf(pages, "頭")).isEqualTo(0);
      assertThat(pages.get(1)).hasSize(2);
    }

    @Test
    @DisplayName("on (the default): the first line is not left alone at the bottom")
    void orphan() throws Exception {
      List<List<Line>> pages = convert(leaving(1).add(p(r("頭" + chars(139)))));

      assertThat(pageOf(pages, "頭")).isEqualTo(1);
      assertThat(pages.get(1)).hasSize(4);
    }

    @Test
    @DisplayName("on: the last line is not left alone at the top, a line moves with it")
    void widow() throws Exception {
      List<List<Line>> pages = convert(leaving(3).add(p(r("頭" + chars(139)))));

      assertThat(pageOf(pages, "頭")).isEqualTo(0);
      assertThat(pages.get(1)).hasSize(2);
    }

    @Test
    @DisplayName("on: a paragraph of three lines is not divided")
    void threeLines() throws Exception {
      List<List<Line>> pages = convert(leaving(2).add(p(r("頭" + chars(99)))));

      assertThat(pageOf(pages, "頭")).isEqualTo(1);
      assertThat(pages.get(1)).hasSize(3);
    }

    @Test
    @DisplayName("keepLines: the whole paragraph moves to the next page")
    void keepLines() throws Exception {
      List<List<Line>> pages =
          convert(leaving(2).add(p("<w:keepLines/>" + NO_WIDOW, r("頭" + chars(139)))));

      assertThat(pageOf(pages, "頭")).isEqualTo(1);
      assertThat(pages.get(1)).hasSize(4);
    }
  }

  @Nested
  @DisplayName("space before at the top of a page")
  class SpaceBeforeAtPageTop {

    private static final String SPACE = "<w:spacing w:before=\"480\"/>";

    @Test
    @DisplayName("it is kept on the first page of the document")
    void firstPage() throws Exception {
      float plain = convert(new DocxBuilder().add(p(r("一")))).get(0).get(0).baseline();
      float spaced = convert(new DocxBuilder().add(p(SPACE, r("一")))).get(0).get(0).baseline();

      assertThat(spaced - plain).isCloseTo(24f, TOL);
    }

    @Test
    @DisplayName("it is ignored on a page reached because the previous page is full")
    void afterAutomaticBreak() throws Exception {
      List<List<Line>> pages = convert(leaving(0).add(p(SPACE, r("次ページ"))));

      assertThat(pages.get(1).get(0).baseline()).isCloseTo(pages.get(0).get(0).baseline(), TOL);
    }

    @Test
    @DisplayName("it is ignored after a page break and with pageBreakBefore")
    void afterExplicitBreak() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().settings(WORD_2013)
          .add(p(r("一") + DocxBuilder.pageBreak())).add(p(SPACE, r("二")))
          .add(p("<w:pageBreakBefore/>" + SPACE, r("三"))));

      assertThat(pages).hasSize(3);
      assertThat(pages.get(1).get(0).baseline()).isCloseTo(pages.get(0).get(0).baseline(), TOL);
      assertThat(pages.get(2).get(0).baseline()).isCloseTo(pages.get(0).get(0).baseline(), TOL);
    }
  }

  @Nested
  @DisplayName("line fitting of justified text")
  class JustifiedFitting {

    private static final String PUNCTUATION =
        "<w:characterSpacingControl w:val=\"compressPunctuation\"/>";
    private static final String BOTH = "<w:jc w:val=\"both\"/>";

    @Test
    @DisplayName("consecutive punctuation is compressed by half a character")
    void consecutivePunctuation() throws Exception {
      Line line = convert(new DocxBuilder().settings(PUNCTUATION)
          .add(p("<w:jc w:val=\"left\"/>", r("あ。）い")))).get(0).get(0);

      assertThat(line.leftOf("い")).isCloseTo(LEFT + 25, TOL);
    }

    @Test
    @DisplayName("punctuation is compressed to put a character which is out by half a character")
    void squeeze() throws Exception {
      // The line is 455pt wide (a right indent of -3.7pt), so the 46th character is out of
      // it by 5pt: half a character, which the two 、 can give.
      String text = chars(20) + "、" + chars(20) + "、" + chars(3) + "末" + chars(10);
      String wide = "<w:ind w:right=\"-74\"/>";
      List<Line> justified = convert(new DocxBuilder().settings(PUNCTUATION)
          .add(p(wide + BOTH, r(text)))).get(0);
      List<Line> left = convert(new DocxBuilder().settings(PUNCTUATION)
          .add(p(wide + "<w:jc w:val=\"left\"/>", r(text)))).get(0);
      List<Line> narrower = convert(new DocxBuilder().settings(PUNCTUATION)
          .add(p(BOTH, r(text)))).get(0);

      assertThat(justified.get(0).text()).endsWith("末");
      assertThat(justified.get(0).right()).isCloseTo(RIGHT + 3.7f, TOL);
      assertThat(left.get(0).text()).doesNotContain("末");
      assertThat(narrower.get(0).text()).doesNotContain("末");
    }

    @Test
    @DisplayName("punctuation which cannot start a line is squeezed into the line")
    void kinsokuSqueeze() throws Exception {
      String text = chars(15) + "、" + chars(15) + "、" + chars(13) + "。" + chars(10);
      List<Line> page = convert(new DocxBuilder().settings(PUNCTUATION)
          .add(p(BOTH, r(text)))).get(0);

      assertThat(page.get(0).text()).endsWith("。");
      assertThat(page.get(0).right()).isCloseTo(RIGHT, TOL);
    }

    @Test
    @DisplayName("without compression, punctuation which cannot start a line hangs out")
    void hanging() throws Exception {
      List<Line> page = convert(new DocxBuilder()
          .add(p("<w:jc w:val=\"left\"/>", r(chars(45) + "。" + chars(10))))).get(0);

      assertThat(page.get(0).text()).endsWith("。");
      assertThat(page.get(0).right()).isGreaterThan(RIGHT + 5);
    }

    @Test
    @DisplayName("overflowPunct off: the character before the punctuation moves down with it")
    void noHanging() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p(
          "<w:overflowPunct w:val=\"0\"/><w:jc w:val=\"left\"/>",
          r(chars(44) + "尾。" + chars(10))))).get(0);

      assertThat(page.get(0).text()).doesNotContain("尾");
      assertThat(page.get(1).text()).startsWith("尾。");
    }

    /**
     * Returns the first line of a justified paragraph of nine words (40 characters, eight
     * spaces) in a line which is narrower than the words by {@code over} points.
     */
    private Line nineWords(float over) throws Exception {
      float space = convert(new DocxBuilder().add(p(r("あ い")))).get(0).get(0).leftOf("い")
          - LEFT - 10;
      float width = 400 + 8 * space - over;
      int rightIndent = Math.round((RIGHT - LEFT - width) * 20);
      String words = "あああああ ".repeat(7) + "ああああ 末 次の行";
      return convert(new DocxBuilder().add(
          p("<w:ind w:right=\"" + rightIndent + "\"/>" + BOTH, r(words)))).get(0).get(0);
    }

    @Test
    @DisplayName("spaces shrink a little to put one more word on a line")
    void shrinkSpaces() throws Exception {
      assertThat(nineWords(2f).text()).endsWith("末");
    }

    @Test
    @DisplayName("spaces do not shrink by more than a quarter of their width")
    void shrinkSpacesLimit() throws Exception {
      assertThat(nineWords(8f).text()).doesNotContain("末");
    }

    @Test
    @DisplayName("a word which fits is on the line as it is")
    void wordFits() throws Exception {
      assertThat(nineWords(-1f).text()).endsWith("末");
    }
  }
}
