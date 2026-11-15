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
import static org.assertj.core.api.Assertions.within;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.assertj.core.data.Offset;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Tests paragraph formatting by the positions of the text in the PDF file.
 *
 * <p>The documents have no styles, so the font size is 10pt, and the test font (Noto Sans JP)
 *     is used for everything: a Japanese character is exactly 10pt wide. The page is A4 with
 *     72pt margins: the text area is from x = 72 to x = 523.3, 45 characters wide.</p>
 */
@DisplayName("WordToPdfUtil: paragraph formatting")
class WordToPdfParagraphTest {

  private static final float LEFT = 72f;
  private static final float RIGHT = 523.3f;
  private static final Offset<Float> TOL = within(0.6f);
  private static final String TAB = "<w:r><w:tab/></w:r>";

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  private List<List<Line>> convert(DocxBuilder docx) throws Exception {
    Path word = docx.write(tempDir.resolve("test.docx"));
    Path pdf = tempDir.resolve("test.pdf");
    @SuppressWarnings("null")
    Path font = Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
    WordToPdfUtil.generate(word, pdf, PdfGenerateOptions.builderForExplicitFont(font).build());
    return PdfLines.of(pdf);
  }

  private static String chars(int count) {
    return "あ".repeat(count);
  }

  /** Returns the distance of the baselines of two plain paragraphs. */
  private float basePitch() throws Exception {
    List<Line> page = convert(new DocxBuilder().add(p(r("一"))).add(p(r("二")))).get(0);
    return page.get(1).baseline() - page.get(0).baseline();
  }

  /** Returns the distance of the baselines of the lines which contain the texts. */
  private static float pitch(List<List<Line>> pages, String upper, String lower) {
    return PdfLines.find(pages, lower).baseline() - PdfLines.find(pages, upper).baseline();
  }

  @Nested
  @DisplayName("indents")
  class Indents {

    @Test
    @DisplayName("left indent moves all the lines")
    void left() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:ind w:left=\"720\"/>",
          r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 36, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT + 36, TOL);
    }

    @Test
    @DisplayName("first line indent moves only the first line")
    void firstLine() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:ind w:firstLine=\"400\"/>",
          r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 20, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT, TOL);
    }

    @Test
    @DisplayName("hanging indent moves the first line to the left of the others")
    void hanging() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(
          p("<w:ind w:left=\"720\" w:hanging=\"360\"/>", r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 18, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT + 36, TOL);
    }

    @Test
    @DisplayName("right indent narrows the lines")
    void right() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:ind w:right=\"1440\"/>",
          r(chars(60))))).get(0);

      // (451.3 - 72) / 10 = 37 characters
      assertThat(page.get(0).text()).hasSize(37);
      assertThat(page.get(0).right()).isLessThanOrEqualTo(RIGHT - 72 + 0.6f);
    }

    @Test
    @DisplayName("indents in characters use the default font size as the unit")
    void inCharacters() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(
          p("<w:ind w:leftChars=\"200\" w:left=\"9999\" w:firstLineChars=\"100\"/>",
              r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 30, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT + 20, TOL);
    }

    @Test
    @DisplayName("left and hanging indents in characters: the hanging one is added to the left")
    void hangingInCharacters() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(
          p("<w:ind w:leftChars=\"100\" w:hangingChars=\"100\"/>", r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 10, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT + 20, TOL);
    }

    @Test
    @DisplayName("a direct indent overrides the indent of the paragraph style")
    void directOverridesStyle() throws Exception {
      String styles = "<w:style w:type=\"paragraph\" w:styleId=\"S\"><w:name w:val=\"S\"/>"
          + "<w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>";
      List<Line> page = convert(new DocxBuilder().styles(styles)
          .add(p("<w:pStyle w:val=\"S\"/>", r("スタイル")))
          .add(p("<w:pStyle w:val=\"S\"/><w:ind w:left=\"1440\"/>", r("直接")))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT + 36, TOL);
      assertThat(page.get(1).left()).isCloseTo(LEFT + 72, TOL);
    }
  }

  @Nested
  @DisplayName("alignment")
  class Alignment {

    @Test
    @DisplayName("center")
    void center() throws Exception {
      Line line = convert(new DocxBuilder().add(p("<w:jc w:val=\"center\"/>", r("あいう"))))
          .get(0).get(0);

      assertThat((line.left() + line.right()) / 2).isCloseTo((LEFT + RIGHT) / 2, TOL);
    }

    @Test
    @DisplayName("right")
    void right() throws Exception {
      Line line = convert(new DocxBuilder().add(p("<w:jc w:val=\"right\"/>", r("あいう"))))
          .get(0).get(0);

      assertThat(line.right()).isCloseTo(RIGHT, TOL);
      assertThat(line.left()).isCloseTo(RIGHT - 30, TOL);
    }

    @Test
    @DisplayName("left: lines are not stretched")
    void left() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:jc w:val=\"left\"/>",
          r(chars(60))))).get(0);

      assertThat(page.get(0).right()).isCloseTo(LEFT + 450, TOL);
    }

    @Test
    @DisplayName("both: lines are stretched to the right edge, except the last line")
    void both() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:jc w:val=\"both\"/>",
          r(chars(60))))).get(0);

      assertThat(page.get(0).left()).isCloseTo(LEFT, TOL);
      assertThat(page.get(0).right()).isCloseTo(RIGHT, TOL);
      assertThat(page.get(1).right()).isCloseTo(LEFT + 150, TOL);
    }

    @Test
    @DisplayName("both: a line ended by a line break is not stretched")
    void bothWithLineBreak() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p("<w:jc w:val=\"both\"/>",
          r(chars(40)) + "<w:r><w:br/></w:r>" + r(chars(60))))).get(0);

      assertThat(page.get(0).right()).isCloseTo(LEFT + 400, TOL);
      assertThat(page.get(1).right()).isCloseTo(RIGHT, TOL);
    }

    @Test
    @DisplayName("distribute: the last line is stretched too")
    void distribute() throws Exception {
      Line line = convert(new DocxBuilder().add(p("<w:jc w:val=\"distribute\"/>",
          r("あいう")))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT, TOL);
      assertThat(line.right()).isCloseTo(RIGHT, TOL);
    }
  }

  @Nested
  @DisplayName("tabs")
  class Tabs {

    @Test
    @DisplayName("without tab stops, a tab moves to the next multiple of 36pt")
    void defaultTabs() throws Exception {
      Line line = convert(new DocxBuilder().add(p(r("あ") + TAB + r("い") + TAB + r("う"))))
          .get(0).get(0);

      assertThat(line.leftOf("い")).isCloseTo(LEFT + 36, TOL);
      assertThat(line.leftOf("う")).isCloseTo(LEFT + 72, TOL);
    }

    @Test
    @DisplayName("defaultTabStop of the settings changes the interval")
    void defaultTabStopSetting() throws Exception {
      Line line = convert(new DocxBuilder().settings("<w:defaultTabStop w:val=\"840\"/>")
          .add(p(r("あ") + TAB + r("い")))).get(0).get(0);

      assertThat(line.leftOf("い")).isCloseTo(LEFT + 42, TOL);
    }

    @Test
    @DisplayName("left tab stop: the text starts at the stop")
    void leftTab() throws Exception {
      Line line = convert(new DocxBuilder().add(
          p("<w:tabs><w:tab w:val=\"left\" w:pos=\"2880\"/></w:tabs>", r("あ") + TAB + r("い"))))
              .get(0).get(0);

      assertThat(line.leftOf("い")).isCloseTo(LEFT + 144, TOL);
    }

    @Test
    @DisplayName("right tab stop: the text ends at the stop")
    void rightTab() throws Exception {
      Line line = convert(new DocxBuilder().add(
          p("<w:tabs><w:tab w:val=\"right\" w:pos=\"6000\"/></w:tabs>",
              r("あ") + TAB + r("右端")))).get(0).get(0);

      assertThat(line.right()).isCloseTo(LEFT + 300, TOL);
    }

    @Test
    @DisplayName("center tab stop: the text is centered at the stop")
    void centerTab() throws Exception {
      Line line = convert(new DocxBuilder().add(
          p("<w:tabs><w:tab w:val=\"center\" w:pos=\"4000\"/></w:tabs>",
              r("あ") + TAB + r("中央")))).get(0).get(0);

      assertThat((line.leftOf("中") + line.right()) / 2).isCloseTo(LEFT + 200, TOL);
    }

    @ParameterizedTest(name = "{0}")
    @CsvSource({"dot, .....", "hyphen, -----", "underscore, _____"})
    @DisplayName("leader: the tab is filled with the leader characters")
    void leader(String leader, String expected) throws Exception {
      Line line = convert(new DocxBuilder().add(
          p("<w:tabs><w:tab w:val=\"right\" w:leader=\"" + leader + "\" w:pos=\"6000\"/></w:tabs>",
              r("見出し") + TAB + r("12")))).get(0).get(0);

      assertThat(line.text()).startsWith("見出し" + expected).endsWith("12");
      assertThat(line.right()).isCloseTo(LEFT + 300, TOL);
    }

    @Test
    @DisplayName("a hanging indent works as a tab stop at the left indent")
    void hangingIndentAsTabStop() throws Exception {
      Line line = convert(new DocxBuilder().add(
          p("<w:ind w:left=\"1000\" w:hanging=\"1000\"/>", r("1") + TAB + r("本文"))))
              .get(0).get(0);

      assertThat(line.leftOf("本")).isCloseTo(LEFT + 50, TOL);
    }

    @Test
    @DisplayName("positional tab to the right: the text ends at the right edge")
    void positionalTab() throws Exception {
      Line line = convert(new DocxBuilder().add(p(r("左")
          + "<w:r><w:ptab w:relativeTo=\"margin\" w:alignment=\"right\" w:leader=\"none\"/></w:r>"
          + r("右")))).get(0).get(0);

      assertThat(line.right()).isCloseTo(RIGHT, TOL);
    }
  }

  @Nested
  @DisplayName("spacing between paragraphs")
  class ParagraphSpacing {

    @Test
    @DisplayName("space before and space after are added between paragraphs")
    void beforeAndAfter() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("一")))
          .add(p("<w:spacing w:before=\"240\" w:after=\"120\"/>", r("二"))).add(p(r("三"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base + 12, TOL);
      assertThat(pitch(pages, "二", "三")).isCloseTo(base + 6, TOL);
    }

    @Test
    @DisplayName("space after of a paragraph and space before of the next are both added")
    void afterPlusBefore() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(new DocxBuilder()
          .add(p("<w:spacing w:after=\"200\"/>", r("一")))
          .add(p("<w:spacing w:before=\"100\"/>", r("二"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base + 15, TOL);
    }

    @Test
    @DisplayName("spacing in lines: a line is 12pt without the document grid")
    void inLines() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("一")))
          .add(p("<w:spacing w:beforeLines=\"100\" w:before=\"9999\"/>", r("二"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base + 12, TOL);
    }

    @Test
    @DisplayName("spacing in lines: a line is the line pitch with the document grid")
    void inLinesWithGrid() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder()
          .addToSectPr("<w:docGrid w:type=\"lines\" w:linePitch=\"480\"/>").add(p(r("一")))
          .add(p("<w:spacing w:beforeLines=\"50\"/>", r("二"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(24 + 12, TOL);
    }

    @Test
    @DisplayName("contextualSpacing: no space between paragraphs of the same style")
    void contextualSpacing() throws Exception {
      final float base = basePitch();
      String spaced = "<w:spacing w:before=\"240\" w:after=\"240\"/>";
      List<List<Line>> pages = convert(new DocxBuilder()
          .add(p(spaced + "<w:contextualSpacing/>", r("一")))
          .add(p(spaced + "<w:contextualSpacing/>", r("二"))).add(p(spaced, r("三")))
          .add(p(r("四"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base, TOL);
      assertThat(pitch(pages, "二", "三")).isCloseTo(base, TOL);
      assertThat(pitch(pages, "三", "四")).isCloseTo(base + 12, TOL);
    }

    @Test
    @DisplayName("a paragraph border takes room: its space and its width")
    void border() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(new DocxBuilder().add(p("<w:pBdr><w:bottom "
          + "w:val=\"single\" w:sz=\"8\" w:space=\"3\" w:color=\"000000\"/></w:pBdr>", r("一")))
          .add(p(r("二"))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base + 4, TOL);
    }
  }

  @Nested
  @DisplayName("line spacing")
  class LineSpacing {

    private float linePitch(String ppr) throws Exception {
      List<Line> page = convert(new DocxBuilder().add(p(ppr, r(chars(100))))).get(0);
      return page.get(1).baseline() - page.get(0).baseline();
    }

    @Test
    @DisplayName("multiple: the line height is multiplied")
    void multiple() throws Exception {
      float base = linePitch("");

      assertThat(linePitch("<w:spacing w:line=\"480\" w:lineRule=\"auto\"/>"))
          .isCloseTo(base * 2, TOL);
      assertThat(linePitch("<w:spacing w:line=\"360\"/>")).isCloseTo(base * 1.5f, TOL);
    }

    @Test
    @DisplayName("exact: the line height is the value")
    void exact() throws Exception {
      assertThat(linePitch("<w:spacing w:line=\"400\" w:lineRule=\"exact\"/>"))
          .isCloseTo(20f, TOL);
      assertThat(linePitch("<w:spacing w:line=\"160\" w:lineRule=\"exact\"/>"))
          .isCloseTo(8f, TOL);
    }

    @Test
    @DisplayName("at least: the larger of the value and the height of the text")
    void atLeast() throws Exception {
      float base = linePitch("");

      assertThat(linePitch("<w:spacing w:line=\"800\" w:lineRule=\"atLeast\"/>"))
          .isCloseTo(40f, TOL);
      assertThat(linePitch("<w:spacing w:line=\"100\" w:lineRule=\"atLeast\"/>"))
          .isCloseTo(base, TOL);
    }

    @Test
    @DisplayName("document grid: lines are placed at the line pitch")
    void grid() throws Exception {
      List<Line> page = convert(new DocxBuilder()
          .addToSectPr("<w:docGrid w:type=\"lines\" w:linePitch=\"480\"/>")
          .add(p(r(chars(100))))).get(0);

      assertThat(page.get(1).baseline() - page.get(0).baseline()).isCloseTo(24f, TOL);
    }

    @Test
    @DisplayName("document grid: text taller than the pitch takes two lines of the grid")
    void gridWithTallText() throws Exception {
      List<Line> page = convert(new DocxBuilder()
          .addToSectPr("<w:docGrid w:type=\"lines\" w:linePitch=\"480\"/>")
          .add(p(r(chars(100), "<w:sz w:val=\"48\"/>")))).get(0);

      assertThat(page.get(1).baseline() - page.get(0).baseline()).isCloseTo(48f, TOL);
    }

    @Test
    @DisplayName("document grid: snapToGrid off makes the paragraph ignore the grid")
    void snapToGridOff() throws Exception {
      float base = linePitch("");
      List<Line> page = convert(new DocxBuilder()
          .addToSectPr("<w:docGrid w:type=\"lines\" w:linePitch=\"480\"/>")
          .add(p("<w:snapToGrid w:val=\"0\"/>", r(chars(100))))).get(0);

      assertThat(page.get(1).baseline() - page.get(0).baseline()).isCloseTo(base, TOL);
    }

    @Test
    @DisplayName("document grid: lines in table cells and in headers do not snap to the grid")
    void gridNotInCellsAndHeaders() throws Exception {
      float base = linePitch("");
      String cell = "<w:tbl><w:tblGrid><w:gridCol w:w=\"9000\"/></w:tblGrid><w:tr><w:tc>"
          + p(r("セ" + chars(100))) + "</w:tc></w:tr></w:tbl>";
      List<List<Line>> pages = convert(new DocxBuilder()
          .settings("<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://"
              + "schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>")
          .header(p(r("ヘ" + chars(100))))
          .addToSectPr("<w:docGrid w:type=\"lines\" w:linePitch=\"480\"/>").add(cell)
          .add(p(r("本" + chars(100)))));
      List<Line> page = pages.get(0);

      int header = page.indexOf(PdfLines.find(pages, "ヘ"));
      int inCell = page.indexOf(PdfLines.find(pages, "セ"));
      int body = page.indexOf(PdfLines.find(pages, "本"));
      assertThat(page.get(header + 1).baseline() - page.get(header).baseline())
          .isCloseTo(base, TOL);
      assertThat(page.get(inCell + 1).baseline() - page.get(inCell).baseline())
          .isCloseTo(base, TOL);
      assertThat(page.get(body + 1).baseline() - page.get(body).baseline())
          .isCloseTo(24f, TOL);
    }
  }

  @Nested
  @DisplayName("number labels")
  class NumberLabels {

    private static String numbering(String levelContent) {
      return "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/>"
          + "<w:numFmt w:val=\"decimal\"/>" + levelContent + "</w:lvl></w:abstractNum>"
          + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
    }

    private static String item(String text) {
      return p("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>", r(text));
    }

    @Test
    @DisplayName("the label is at the hanging position, and the text at the left indent")
    void labelAndText() throws Exception {
      Line line = convert(new DocxBuilder().numbering(numbering("<w:lvlText w:val=\"%1.\"/>"
          + "<w:pPr><w:ind w:left=\"1000\" w:hanging=\"600\"/></w:pPr>"))
          .add(item("本文" + chars(60)))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT + 20, TOL);
      assertThat(line.leftOf("本")).isCloseTo(LEFT + 50, TOL);
    }

    @Test
    @DisplayName("suffix nothing: the text follows the label immediately")
    void suffixNothing() throws Exception {
      Line line = convert(new DocxBuilder().numbering(numbering("<w:suff w:val=\"nothing\"/>"
          + "<w:lvlText w:val=\"%1)\"/>")).add(item("本文"))).get(0).get(0);

      assertThat(line.leftOf("本")).isCloseTo(line.rightOf(")"), TOL);
    }

    @Test
    @DisplayName("right-justified label: the label ends at the hanging position")
    void rightJustified() throws Exception {
      Line line = convert(new DocxBuilder().numbering(numbering("<w:lvlText w:val=\"%1.\"/>"
          + "<w:lvlJc w:val=\"right\"/><w:pPr><w:ind w:left=\"1000\" w:hanging=\"400\"/></w:pPr>"))
          .add(item("本文"))).get(0).get(0);

      assertThat(line.rightOf(".")).isCloseTo(LEFT + 30, TOL);
      assertThat(line.leftOf("本")).isCloseTo(LEFT + 50, TOL);
    }

    @Test
    @DisplayName("the indent of the paragraph overrides the indent of the numbering level")
    void directIndentOverridesLevel() throws Exception {
      Line line = convert(new DocxBuilder().numbering(numbering("<w:lvlText w:val=\"%1.\"/>"
          + "<w:pPr><w:ind w:left=\"1000\" w:hanging=\"600\"/></w:pPr>"))
          .add(p("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
              + "<w:ind w:left=\"2000\" w:hanging=\"600\"/>", r("本文")))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT + 70, TOL);
      assertThat(line.leftOf("本")).isCloseTo(LEFT + 100, TOL);
    }
  }
}
