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
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

/** Tests character formatting, the precedence of styles, and fonts. */
@DisplayName("WordToPdfUtil: character formatting and styles")
class WordToPdfCharacterTest extends AbstractWordToPdfTest {

  /** Returns the distance of the first two characters of the first line. */
  private float charPitch(DocxBuilder docx) throws Exception {
    Line line = convert(docx).get(0).get(0);
    return line.glyphs().get(1).getXDirAdj() - line.glyphs().get(0).getXDirAdj();
  }

  @Nested
  @DisplayName("size, weight and spacing")
  class SizeWeightSpacing {

    @Test
    @DisplayName("font size")
    void size() throws Exception {
      assertThat(charPitch(new DocxBuilder().add(p(r("ああ", "<w:sz w:val=\"40\"/>")))))
          .isCloseTo(20f, TOL);
    }

    @Test
    @DisplayName("bold without a bold font: emboldened, advancing 2% of the size more")
    void syntheticBold() throws Exception {
      assertThat(charPitch(new DocxBuilder().add(p(r("ああ", "<w:b/><w:sz w:val=\"100\"/>")))))
          .isCloseTo(51f, TOL);
    }

    @SuppressWarnings("null")
    @Test
    @DisplayName("bold with a bold font: the bold font is used, with its own advance")
    void boldFont() throws Exception {
      Path word = new DocxBuilder().add(p(r("ああ", "<w:b/><w:sz w:val=\"100\"/>")))
          .write(tempDir.resolve("bold.docx"));
      Path pdf = tempDir.resolve("bold.pdf");
      Path dir = Path.of(getClass().getResource("/fonts/NotoSansJP").toURI());
      WordToPdfUtil.generate(word, pdf,
          PdfGenerateOptions.builderForExplicitFont(dir.resolve("NotoSansJP-Regular.ttf"))
              .addBoldFontPath(dir.resolve("NotoSansJP-Bold.ttf")).build());

      Line line = PdfLines.of(pdf).get(0).get(0);
      assertThat(line.glyphs().get(1).getXDirAdj() - line.glyphs().get(0).getXDirAdj())
          .isCloseTo(50f, TOL);
      assertThat(line.glyphs().get(0).getFont().getName()).contains("Bold");
    }

    @Test
    @DisplayName("character spacing: expanded and condensed")
    void spacing() throws Exception {
      assertThat(charPitch(new DocxBuilder().add(p(r("ああ", "<w:spacing w:val=\"40\"/>")))))
          .isCloseTo(12f, TOL);
      assertThat(charPitch(new DocxBuilder().add(p(r("ああ", "<w:spacing w:val=\"-40\"/>")))))
          .isCloseTo(8f, TOL);
    }

    @Test
    @DisplayName("superscript and subscript: smaller, and moved from the baseline")
    void superscriptAndSubscript() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("基", "<w:sz w:val=\"40\"/>")
          + r("上", "<w:sz w:val=\"40\"/><w:vertAlign w:val=\"superscript\"/>")
          + r("下", "<w:sz w:val=\"40\"/><w:vertAlign w:val=\"subscript\"/>"))));
      final Line base = PdfLines.find(pages, "基");
      final Line sup = PdfLines.find(pages, "上");
      final Line sub = PdfLines.find(pages, "下");

      assertThat(base.baseline() - sup.baseline()).isCloseTo(20 * 0.215f, TOL);
      assertThat(sub.baseline() - base.baseline()).isCloseTo(20 * 0.14f, TOL);
      assertThat(sup.glyphs().get(0).getWidthDirAdj()).isCloseTo(13f, TOL);
    }

    @Test
    @DisplayName("all caps")
    void caps() throws Exception {
      assertThat(convert(new DocxBuilder().add(p(r("Word file", "<w:caps/>")))).get(0).get(0)
          .text()).isEqualTo("WORDFILE");
    }
  }

  @Nested
  @DisplayName("colors and lines")
  class ColorsAndLines {

    private static final String BIG = "<w:sz w:val=\"80\"/>";

    /** Returns the color at the offset from the start of the baseline of the first line. */
    private int colorAt(String rpr, String text, float dx, float dy) throws Exception {
      Path pdf = generate(new DocxBuilder().add(p(r("基", BIG) + r(text, BIG + rpr))));
      Line line = PdfLines.find(PdfLines.of(pdf), "基");
      return pixel(pdf, LEFT + 40 + dx, line.baseline() + dy);
    }

    @Test
    @DisplayName("text color")
    void color() throws Exception {
      assertThat(colorAt("<w:color w:val=\"FF0000\"/>", "■", 20, -14)).isEqualTo(0xFF0000);
      assertThat(colorAt("", "■", 20, -14)).isEqualTo(0x000000);
      assertThat(colorAt("<w:color w:val=\"auto\"/>", "■", 20, -14)).isEqualTo(0x000000);
    }

    @Test
    @DisplayName("highlight and character shading fill the background of the text")
    void highlightAndShading() throws Exception {
      assertThat(colorAt("<w:highlight w:val=\"yellow\"/>", "　　", 40, -14))
          .isEqualTo(0xFFFF00);
      assertThat(colorAt("<w:shd w:val=\"clear\" w:fill=\"00FF00\"/>", "　　", 40, -14))
          .isEqualTo(0x00FF00);
      assertThat(colorAt("", "　　", 40, -14)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("underline is drawn below the baseline, also under spaces")
    void underline() throws Exception {
      assertThat(dark(colorAt("<w:u w:val=\"single\"/>", "　　", 40, 4.8f))).isTrue();
      assertThat(dark(colorAt("<w:u w:val=\"double\"/>", "　　", 40, 4.8f))).isTrue();
      assertThat(dark(colorAt("<w:u w:val=\"none\"/>", "　　", 40, 4.8f))).isFalse();
      assertThat(dark(colorAt("", "　　", 40, 4.8f))).isFalse();
    }

    @Test
    @DisplayName("strikethrough is drawn through the text")
    void strike() throws Exception {
      assertThat(dark(colorAt("<w:strike/>", "　　", 40, -11.2f))).isTrue();
      assertThat(dark(colorAt("<w:dstrike/>", "　　", 40, -11.2f))).isTrue();
      assertThat(dark(colorAt("", "　　", 40, -11.2f))).isFalse();
    }

    @Test
    @DisplayName("paragraph shading and borders are drawn")
    void paragraphShadingAndBorders() throws Exception {
      Path pdf = generate(new DocxBuilder().add(p("<w:pBdr><w:left w:val=\"single\" w:sz=\"48\" "
          + "w:space=\"4\" w:color=\"0000FF\"/></w:pBdr><w:shd w:val=\"clear\" w:fill=\"FFCC00\"/>",
          r("　　", BIG))));

      assertThat(pixel(pdf, LEFT + 200, 72 + 20)).isEqualTo(0xFFCC00);
      assertThat(pixel(pdf, LEFT - 4 - 3, 72 + 20)).isEqualTo(0x0000FF);
    }
  }

  @Nested
  @DisplayName("precedence of styles")
  class StylePrecedence {

    private static final String STYLES = "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"24\"/>"
        + "</w:rPr></w:rPrDefault></w:docDefaults>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"S\"><w:name w:val=\"S\"/>"
        + "<w:rPr><w:b/><w:sz w:val=\"100\"/></w:rPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"D\"><w:name w:val=\"D\"/>"
        + "<w:basedOn w:val=\"S\"/><w:pPr><w:jc w:val=\"right\"/></w:pPr>"
        + "<w:rPr><w:b w:val=\"0\"/></w:rPr></w:style>"
        + "<w:style w:type=\"character\" w:styleId=\"C\"><w:name w:val=\"C\"/>"
        + "<w:rPr><w:sz w:val=\"40\"/></w:rPr></w:style>";

    @Test
    @DisplayName("document defaults apply to a paragraph without a style")
    void documentDefaults() throws Exception {
      assertThat(charPitch(new DocxBuilder().styles(STYLES).add(p(r("ああ")))))
          .isCloseTo(12f, TOL);
    }

    @Test
    @DisplayName("a paragraph style overrides the document defaults")
    void paragraphStyle() throws Exception {
      assertThat(charPitch(new DocxBuilder().styles(STYLES)
          .add(p("<w:pStyle w:val=\"S\"/>", r("ああ"))))).isCloseTo(51f, TOL);
    }

    @Test
    @DisplayName("a derived style inherits, and overrides what it specifies (bold off)")
    void basedOn() throws Exception {
      Line line = convert(new DocxBuilder().styles(STYLES)
          .add(p("<w:pStyle w:val=\"D\"/>", r("ああ")))).get(0).get(0);

      assertThat(line.glyphs().get(1).getXDirAdj() - line.glyphs().get(0).getXDirAdj())
          .isCloseTo(50f, TOL);
      assertThat(line.right()).isCloseTo(RIGHT, TOL);
    }

    @Test
    @DisplayName("a character style overrides the paragraph style, direct formatting both")
    void characterStyleAndDirect() throws Exception {
      assertThat(charPitch(new DocxBuilder().styles(STYLES)
          .add(p("<w:pStyle w:val=\"D\"/>", r("ああ", "<w:rStyle w:val=\"C\"/>")))))
              .isCloseTo(20f, TOL);
      assertThat(charPitch(new DocxBuilder().styles(STYLES).add(p("<w:pStyle w:val=\"D\"/>",
          r("ああ", "<w:rStyle w:val=\"C\"/><w:sz w:val=\"60\"/>"))))).isCloseTo(30f, TOL);
    }

    @Test
    @DisplayName("the default paragraph style applies to a paragraph without a style")
    void defaultParagraphStyle() throws Exception {
      String styles = "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"N\">"
          + "<w:name w:val=\"Normal\"/><w:rPr><w:sz w:val=\"60\"/></w:rPr></w:style>";

      assertThat(charPitch(new DocxBuilder().styles(styles).add(p(r("ああ")))))
          .isCloseTo(30f, TOL);
    }

    @Test
    @DisplayName("styles based on each other in a cycle do not hang")
    void cycle() throws Exception {
      String styles = "<w:style w:type=\"paragraph\" w:styleId=\"A\"><w:basedOn w:val=\"B\"/>"
          + "<w:rPr><w:sz w:val=\"60\"/></w:rPr></w:style>"
          + "<w:style w:type=\"paragraph\" w:styleId=\"B\"><w:basedOn w:val=\"A\"/></w:style>";

      assertThat(charPitch(new DocxBuilder().styles(styles)
          .add(p("<w:pStyle w:val=\"A\"/>", r("ああ"))))).isCloseTo(30f, TOL);
    }

    @Test
    @DisplayName("a style which does not exist is ignored")
    void missingStyle() throws Exception {
      assertThat(charPitch(new DocxBuilder().add(p("<w:pStyle w:val=\"None\"/>",
          r("ああ", "<w:rStyle w:val=\"None\"/>"))))).isCloseTo(10f, TOL);
    }
  }

  @Nested
  @DisplayName("fonts")
  class Fonts {

    private String fontNames(Path pdf) throws Exception {
      StringBuilder sb = new StringBuilder();
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        for (var name : doc.getPage(0).getResources().getFontNames()) {
          sb.append(doc.getPage(0).getResources().getFont(name).getName()).append(' ');
        }
      }
      return sb.toString();
    }

    @Test
    @DisplayName("explicit font: the font is used whatever fonts the document specifies")
    void explicitFont() throws Exception {
      Path pdf = generate(new DocxBuilder().add(p(r("Arial指定",
          "<w:rFonts w:ascii=\"Arial\" w:eastAsia=\"ＭＳ 明朝\" w:hAnsi=\"Arial\"/>"))));

      assertThat(fontNames(pdf)).contains("NotoSansJP").doesNotContain("Arial");
    }

    @SuppressWarnings("null")
    @Test
    @DisplayName("system fonts with a fallback font: fonts not installed are replaced by it")
    void systemFontsWithFallback() throws Exception {
      Path word = new DocxBuilder().add(p(r("未知のフォント",
          "<w:rFonts w:ascii=\"NoSuchFontXyz\" w:eastAsia=\"NoSuchFontXyz\"/>")))
          .write(tempDir.resolve("system.docx"));
      Path pdf = tempDir.resolve("system.pdf");
      Path font =
          Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
      WordToPdfUtil.generate(word, pdf,
          PdfGenerateOptions.builderForSystemFonts().addRegularFontPath(font).build());

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("未知のフォント");
      }
    }
  }
}
