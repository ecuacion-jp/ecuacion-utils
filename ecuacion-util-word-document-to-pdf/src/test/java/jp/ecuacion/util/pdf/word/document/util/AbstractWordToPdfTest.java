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
import static org.assertj.core.api.Assertions.within;
import java.awt.image.BufferedImage;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.assertj.core.data.Offset;
import org.junit.jupiter.api.io.TempDir;

/**
 * Provides the common parts of the tests which verify the layout by positions in the PDF file.
 *
 * <p>The documents have no styles, so the font size is 10pt, and the test font (Noto Sans JP)
 *     is used for everything: a Japanese character is exactly 10pt wide. The page is A4 with
 *     72pt margins: the text area is from x = 72 to x = 523.3 (45 characters), and from
 *     y = 72 to y = 769.9.</p>
 */
abstract class AbstractWordToPdfTest {

  static final float LEFT = 72f;
  static final float RIGHT = 523.3f;
  static final float BODY_HEIGHT = 697.9f;
  static final Offset<Float> TOL = within(0.6f);
  /** Settings of a document written by Word 2013 or later. */
  static final String WORD_2013 = "<w:compat><w:compatSetting w:name=\"compatibilityMode\" "
      + "w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>";

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  Path generate(DocxBuilder docx) throws Exception {
    Path word = docx.write(tempDir.resolve("test.docx"));
    Path pdf = tempDir.resolve("test.pdf");
    @SuppressWarnings("null")
    Path font =
        Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
    WordToPdfUtil.generate(word, pdf, PdfGenerateOptions.builderForExplicitFont(font).build());
    return pdf;
  }

  List<List<Line>> convert(DocxBuilder docx) throws Exception {
    return PdfLines.of(generate(docx));
  }

  static String chars(int count) {
    return "あ".repeat(count);
  }

  /** Returns the distance of the baselines of two plain paragraphs. */
  float basePitch() throws Exception {
    List<Line> page = convert(new DocxBuilder().add(p(r("一"))).add(p(r("二")))).get(0);
    return page.get(1).baseline() - page.get(0).baseline();
  }

  /** Returns the distance of the baselines of the lines which contain the texts. */
  static float pitch(List<List<Line>> pages, String upper, String lower) {
    return PdfLines.find(pages, lower).baseline() - PdfLines.find(pages, upper).baseline();
  }

  /** Returns the index of the page which has the text, or -1. */
  static int pageOf(List<List<Line>> pages, String text) {
    for (int i = 0; i < pages.size(); i++) {
      for (Line line : pages.get(i)) {
        if (line.text().contains(text)) {
          return i;
        }
      }
    }
    return -1;
  }

  /** Returns whether the color is dark (a part of text or of a line). */
  static boolean dark(int rgb) {
    return ((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF) < 3 * 0x60;
  }

  /** Returns the RGB color of the first page at the position in points from the top-left. */
  static int pixel(Path pdf, float x, float y) throws Exception {
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 144);
      return image.getRGB(Math.round(x * 2), Math.round(y * 2)) & 0xFFFFFF;
    }
  }
}
