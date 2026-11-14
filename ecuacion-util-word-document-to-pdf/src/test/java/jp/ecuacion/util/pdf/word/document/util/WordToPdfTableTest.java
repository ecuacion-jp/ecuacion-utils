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
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Tests tables. The documents are in the compatibility mode of Word 2013 or later, where the
 * left edge of a table is at the left margin (x = 72), and the default cell margins are 5.4pt
 * on the left and the right.
 */
@DisplayName("WordToPdfUtil: tables")
class WordToPdfTableTest extends AbstractWordToPdfTest {

  private static final float MARGIN = 5.4f;

  private static String borders(String tag, String style, int size) {
    StringBuilder sb = new StringBuilder("<w:" + tag + ">");
    for (String edge : new String[] {"top", "left", "bottom", "right", "insideH", "insideV"}) {
      if (tag.equals("tcBorders") && edge.startsWith("inside")) {
        continue;
      }
      sb.append("<w:").append(edge).append(" w:val=\"").append(style).append("\" w:sz=\"")
          .append(size).append("\" w:space=\"0\" w:color=\"000000\"/>");
    }
    return sb.append("</w:").append(tag).append(">").toString();
  }

  private static String cell(String tcpr, String content) {
    return "<w:tc><w:tcPr>" + tcpr + "</w:tcPr>"
        + (content.startsWith("<") ? content : p(r(content))) + "</w:tc>";
  }

  private static String row(String trpr, String... cells) {
    return "<w:tr><w:trPr>" + trpr + "</w:trPr>" + String.join("", cells) + "</w:tr>";
  }

  /** Returns a table whose grid columns have the widths in twips. */
  private static String table(String tblpr, int[] grid, String... rows) {
    StringBuilder sb = new StringBuilder("<w:tbl><w:tblPr>" + tblpr + "</w:tblPr><w:tblGrid>");
    for (int width : grid) {
      sb.append("<w:gridCol w:w=\"").append(width).append("\"/>");
    }
    return sb.append("</w:tblGrid>").append(String.join("", rows)).append("</w:tbl>").toString();
  }

  private static DocxBuilder docx() {
    return new DocxBuilder().settings(WORD_2013);
  }

  @Nested
  @DisplayName("position and widths")
  class PositionAndWidths {

    @Test
    @DisplayName("cells start at the grid columns, the text inside the cell margin")
    void columns() throws Exception {
      Line line = convert(docx().add(table("", new int[] {2000, 3000},
          row("", cell("", "甲"), cell("", "乙"))))).get(0).get(0);

      assertThat(line.leftOf("甲")).isCloseTo(LEFT + MARGIN, TOL);
      assertThat(line.leftOf("乙")).isCloseTo(LEFT + 100 + MARGIN, TOL);
    }

    @Test
    @DisplayName("before Word 2013 the text of the first column is aligned to the margin")
    void oldCompatibilityMode() throws Exception {
      Line line = convert(new DocxBuilder().add(table("", new int[] {2000, 3000},
          row("", cell("", "甲"), cell("", "乙"))))).get(0).get(0);

      assertThat(line.leftOf("甲")).isCloseTo(LEFT, TOL);
    }

    @ParameterizedTest(name = "{0}")
    @CsvSource({"center, 197.65", "right, 323.3"})
    @DisplayName("table alignment")
    void alignment(String jc, float expectedLeft) throws Exception {
      Line line = convert(docx().add(table("<w:jc w:val=\"" + jc + "\"/>", new int[] {4000},
          row("", cell("", "甲"))))).get(0).get(0);

      assertThat(line.left()).isCloseTo(expectedLeft + MARGIN, TOL);
    }

    @Test
    @DisplayName("table indent")
    void indent() throws Exception {
      Line line = convert(docx().add(table("<w:tblInd w:w=\"1000\" w:type=\"dxa\"/>",
          new int[] {4000}, row("", cell("", "甲"))))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT + 50 + MARGIN, TOL);
    }

    @Test
    @DisplayName("width in percent: the columns are scaled to the ratio of the text width")
    void percentWidth() throws Exception {
      Line line = convert(docx().add(table("<w:tblW w:w=\"2500\" w:type=\"pct\"/>",
          new int[] {1000, 1000}, row("", cell("", "甲"), cell("", "乙"))))).get(0).get(0);

      assertThat(line.leftOf("乙")).isCloseTo(LEFT + (RIGHT - LEFT) / 4 + MARGIN, TOL);
    }

    @Test
    @DisplayName("cell margins of the table and of a cell")
    void cellMargins() throws Exception {
      String tableMargins = "<w:tblCellMar><w:left w:w=\"400\" w:type=\"dxa\"/></w:tblCellMar>";
      String cellMargins = "<w:tcMar><w:left w:w=\"800\" w:type=\"dxa\"/></w:tcMar>";
      Line line = convert(docx().add(table(tableMargins, new int[] {3000, 3000},
          row("", cell("", "甲"), cell(cellMargins, "乙"))))).get(0).get(0);

      assertThat(line.leftOf("甲")).isCloseTo(LEFT + 20, TOL);
      assertThat(line.leftOf("乙")).isCloseTo(LEFT + 150 + 40, TOL);
    }

    @Test
    @DisplayName("text wraps in the width of the cell")
    void wrap() throws Exception {
      List<Line> page = convert(docx().add(table("", new int[] {2216},
          row("", cell("", chars(25)))))).get(0);

      // 110.8 - 5.4 * 2 = 100pt: 10 characters a line
      assertThat(page).hasSize(3);
      assertThat(page.get(0).text()).hasSize(10);
    }

    @Test
    @DisplayName("a column narrower than a word is widened, taking width from the others")
    void widenedForLongWord() throws Exception {
      float wordWidth = convert(docx().add(p(r("Supercalifragilistic")))).get(0).get(0).right()
          - LEFT;
      List<List<Line>> pages = convert(docx().add(table("", new int[] {1000, 5000},
          row("", cell("", "Supercalifragilistic"), cell("", "乙")))));

      Line line = pages.get(0).get(0);
      assertThat(line.text()).isEqualTo("Supercalifragilistic乙");
      assertThat(line.leftOf("乙")).isCloseTo(LEFT + wordWidth + MARGIN * 3, TOL);
    }

    @Test
    @DisplayName("fixed layout: the column is not widened and the word is broken")
    void fixedLayout() throws Exception {
      List<Line> page = convert(docx().add(table("<w:tblLayout w:type=\"fixed\"/>",
          new int[] {1000, 5000}, row("", cell("", "Supercalifragilistic"), cell("", "乙")))))
              .get(0);

      assertThat(page.size()).isGreaterThan(1);
      assertThat(PdfLines.find(List.of(page), "乙").leftOf("乙"))
          .isCloseTo(LEFT + 50 + MARGIN, TOL);
    }
  }

  @Nested
  @DisplayName("merged cells")
  class MergedCells {

    @Test
    @DisplayName("a horizontally merged cell spans grid columns")
    void gridSpan() throws Exception {
      List<Line> page = convert(docx().add(table("", new int[] {2000, 2000, 2000},
          row("", cell("<w:gridSpan w:val=\"2\"/>", "結合"), cell("", "右")),
          row("", cell("", "甲"), cell("", "乙"), cell("", "丙"))))).get(0);

      assertThat(page.get(0).leftOf("右")).isCloseTo(LEFT + 200 + MARGIN, TOL);
      assertThat(page.get(1).leftOf("丙")).isCloseTo(LEFT + 200 + MARGIN, TOL);
    }

    @Test
    @DisplayName("a vertically merged cell: centered content is in the middle of its rows")
    void verticalMerge() throws Exception {
      List<List<Line>> pages = convert(docx().add(table("", new int[] {2000, 2000},
          row("", cell("<w:vMerge w:val=\"restart\"/><w:vAlign w:val=\"center\"/>", "結合"),
              cell("", "一")),
          row("", cell("<w:vMerge/>", ""), cell("", "二")),
          row("", cell("<w:vMerge/>", ""), cell("", "三")))));

      assertThat(PdfLines.find(pages, "結合").baseline())
          .isCloseTo(PdfLines.find(pages, "二").baseline(), TOL);
    }
  }

  @Nested
  @DisplayName("row heights and vertical alignment")
  class RowHeights {

    @Test
    @DisplayName("minimum row height")
    void atLeast() throws Exception {
      List<List<Line>> pages = convert(docx().add(table("", new int[] {3000},
          row("<w:trHeight w:val=\"1000\"/>", cell("", "一")), row("", cell("", "二")))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(50f, TOL);
    }

    @Test
    @DisplayName("minimum row height: a taller content makes the row taller")
    void atLeastWithTallContent() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(docx().add(table("", new int[] {3000},
          row("<w:trHeight w:val=\"100\"/>", cell("", p(r("一")) + p(r("中")))),
          row("", cell("", "二")))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(base * 2, TOL);
    }

    @Test
    @DisplayName("exact row height")
    void exact() throws Exception {
      List<List<Line>> pages = convert(docx().add(table("", new int[] {3000},
          row("<w:trHeight w:val=\"600\" w:hRule=\"exact\"/>", cell("", "一")),
          row("", cell("", "二")))));

      assertThat(pitch(pages, "一", "二")).isCloseTo(30f, TOL);
    }

    @Test
    @DisplayName("vertical alignment: center and bottom")
    void verticalAlignment() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(docx().add(table("", new int[] {2000, 2000, 2000},
          row("<w:trHeight w:val=\"1200\"/>", cell("", "上"),
              cell("<w:vAlign w:val=\"center\"/>", "中"),
              cell("<w:vAlign w:val=\"bottom\"/>", "下")))));
      final Line top = PdfLines.find(pages, "上");
      final Line center = PdfLines.find(pages, "中");
      final Line bottom = PdfLines.find(pages, "下");

      assertThat(center.baseline() - top.baseline()).isCloseTo((60 - base) / 2, TOL);
      assertThat(bottom.baseline() - top.baseline()).isCloseTo(60 - base, TOL);
    }

    @Test
    @DisplayName("the top margin of a row is the largest top margin of its cells")
    void rowTopMargin() throws Exception {
      Line line = convert(docx().add(table("", new int[] {2000, 2000}, row("", cell("", "甲"),
          cell("<w:tcMar><w:top w:w=\"400\" w:type=\"dxa\"/></w:tcMar>", "乙"))))).get(0)
              .get(0);

      assertThat(line.text()).isEqualTo("甲乙");
    }
  }

  @Nested
  @DisplayName("borders and shading")
  class BordersAndShading {

    private float rowPitch(String tblpr) throws Exception {
      return pitch(convert(docx().add(table(tblpr, new int[] {3000}, row("", cell("", "一")),
          row("", cell("", "二"))))), "一", "二");
    }

    @Test
    @DisplayName("a horizontal border adds its width to the height of a row")
    void borderRoom() throws Exception {
      float base = rowPitch("");

      assertThat(base).isCloseTo(basePitch(), TOL);
      assertThat(rowPitch(borders("tblBorders", "single", 4)) - base).isCloseTo(0.5f, TOL);
      assertThat(rowPitch(borders("tblBorders", "single", 48)) - base).isCloseTo(6f, TOL);
    }

    @Test
    @DisplayName("a double border takes three times its width")
    void doubleBorderRoom() throws Exception {
      assertThat(rowPitch(borders("tblBorders", "double", 16)) - rowPitch(""))
          .isCloseTo(6f, TOL);
    }

    @Test
    @DisplayName("the text moves right by half of the left border")
    void leftBorder() throws Exception {
      Line line = convert(docx().add(table(borders("tblBorders", "single", 48),
          new int[] {3000}, row("", cell("", "甲"))))).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT + MARGIN + 3, TOL);
    }

    @Test
    @DisplayName("borders are drawn: of the table, and of a cell overriding the table")
    void bordersDrawn() throws Exception {
      Path pdf = generate(docx().add(table(borders("tblBorders", "single", 32),
          new int[] {2000, 2000},
          row("<w:trHeight w:val=\"800\"/>", cell("", "甲"),
              cell("<w:tcBorders><w:right w:val=\"nil\"/></w:tcBorders>", "乙")))));

      // the left border of the table (4pt wide at x = 72) and the vertical border inside
      assertThat(pixel(pdf, LEFT, 90)).isEqualTo(0x000000);
      assertThat(pixel(pdf, LEFT + 100, 90)).isEqualTo(0x000000);
      // the right border is removed by the cell
      assertThat(pixel(pdf, LEFT + 200, 90)).isEqualTo(0xFFFFFF);
      // inside a cell
      assertThat(pixel(pdf, LEFT + 150, 100)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("a table without borders draws no lines")
    void noBorders() throws Exception {
      Path pdf = generate(docx().add(table("", new int[] {2000, 2000},
          row("<w:trHeight w:val=\"800\"/>", cell("", "甲"), cell("", "乙")))));

      assertThat(pixel(pdf, LEFT, 90)).isEqualTo(0xFFFFFF);
      assertThat(pixel(pdf, LEFT + 100, 90)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("cell shading fills the cell")
    void shading() throws Exception {
      Path pdf = generate(docx().add(table("", new int[] {2000, 2000},
          row("<w:trHeight w:val=\"800\"/>",
              cell("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FF0000\"/>", ""),
              cell("", "")))));

      assertThat(pixel(pdf, LEFT + 50, 92)).isEqualTo(0xFF0000);
      assertThat(pixel(pdf, LEFT + 150, 92)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("table style: the first row and banded rows are shaded as tblLook says")
    void tableStyleConditions() throws Exception {
      String styles = "<w:style w:type=\"table\" w:styleId=\"T\"><w:name w:val=\"T\"/>"
          + "<w:tblStylePr w:type=\"firstRow\"><w:tcPr><w:shd w:val=\"clear\" w:fill=\"0000FF\"/>"
          + "</w:tcPr></w:tblStylePr><w:tblStylePr w:type=\"band1Horz\"><w:tcPr>"
          + "<w:shd w:val=\"clear\" w:fill=\"00FF00\"/></w:tcPr></w:tblStylePr></w:style>";
      String height = "<w:trHeight w:val=\"400\" w:hRule=\"exact\"/>";
      Path pdf = generate(docx().styles(styles).add(table(
          "<w:tblStyle w:val=\"T\"/><w:tblLook w:firstRow=\"1\" w:noHBand=\"0\"/>",
          new int[] {2000}, row(height, cell("", "")), row(height, cell("", "")),
          row(height, cell("", "")), row(height, cell("", "")))));

      assertThat(pixel(pdf, LEFT + 50, 72 + 10)).isEqualTo(0x0000FF);
      assertThat(pixel(pdf, LEFT + 50, 72 + 30)).isEqualTo(0x00FF00);
      assertThat(pixel(pdf, LEFT + 50, 72 + 50)).isEqualTo(0xFFFFFF);
      assertThat(pixel(pdf, LEFT + 50, 72 + 70)).isEqualTo(0x00FF00);
    }
  }

  @Nested
  @DisplayName("content of cells")
  class CellContent {

    @Test
    @DisplayName("paragraph alignment and indent work in a cell")
    void paragraphFormatting() throws Exception {
      List<List<Line>> pages = convert(docx().add(table("", new int[] {4000},
          row("", cell("", p("<w:jc w:val=\"right\"/>", r("右"))
              + p("<w:ind w:left=\"400\"/>", r("字下げ")))))));

      assertThat(PdfLines.find(pages, "右").right()).isCloseTo(LEFT + 200 - MARGIN, TOL);
      assertThat(PdfLines.find(pages, "字下げ").left()).isCloseTo(LEFT + MARGIN + 20, TOL);
    }

    @Test
    @DisplayName("a nested table is laid out in the cell")
    void nested() throws Exception {
      String inner = table("", new int[] {1000, 1000}, row("", cell("", "内一"),
          cell("", "内二")));
      List<List<Line>> pages = convert(docx().add(table("", new int[] {2000, 4000},
          row("", cell("", "外"), cell("", p(r("前")) + inner + p(r("後")))))));

      Line innerLine = PdfLines.find(pages, "内一");
      assertThat(innerLine.leftOf("内一")).isCloseTo(LEFT + 100 + MARGIN + MARGIN, TOL);
      assertThat(innerLine.leftOf("内二")).isCloseTo(LEFT + 100 + MARGIN + 50 + MARGIN, TOL);
      assertThat(PdfLines.find(pages, "前").baseline()).isLessThan(innerLine.baseline());
      assertThat(PdfLines.find(pages, "後").baseline()).isGreaterThan(innerLine.baseline());
    }

    @Test
    @DisplayName("the paragraph after a table follows it")
    void paragraphAfterTable() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(docx().add(table("", new int[] {3000},
          row("", cell("", "表")))).add(p(r("後"))));

      assertThat(pitch(pages, "表", "後")).isCloseTo(base, TOL);
    }
  }
}
