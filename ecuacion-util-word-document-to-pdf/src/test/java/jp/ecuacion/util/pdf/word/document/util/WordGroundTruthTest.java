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

import static org.assertj.core.api.Assertions.assertThat;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;
import jp.ecuacion.util.pdf.common.internal.SystemFontLocator;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

/**
 * Compares the output with PDF files exported by Microsoft Word (the ground truth).
 *
 * <p>Each {@code ground-truth/<name>.docx} has {@code <name>.docx.pdf} exported from it by Word
 *     for Mac (see {@code src/test/ground-truth-tools}). The Word file is converted with the
 *     fonts of the system, and every line of the result must have the same text, at the same
 *     position, as the line of Word's PDF file. So both where lines and pages break and the
 *     heights of lines are verified.</p>
 *
 * <p>The tests are skipped where the fonts used by the files are not installed.</p>
 */
@DisplayName("comparison with PDF files exported by Word")
class WordGroundTruthTest {

  /** The fonts used by the ground truth files. */
  private static final List<String> FONTS = List.of("游明朝", "游ゴシック", "ＭＳ 明朝", "メイリオ",
      "Century", "Arial", "Times New Roman");

  /** The allowed difference of the position of a line in points. */
  private static final float TOLERANCE = 1.0f;

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  /** A line of text: the baseline, the left edge and the normalized text. */
  private record Line(float baseline, float left, String text) {
  }

  /** Returns the names of the Word files in the ground-truth directory. */
  @SuppressWarnings("null")
  static List<String> names() throws Exception {
    URL dir = WordGroundTruthTest.class.getResource("/ground-truth");
    if (dir == null) {
      return List.of();
    }
    try (Stream<Path> files = Files.list(Path.of(dir.toURI()))) {
      return files.map(f -> f.getFileName().toString()).filter(n -> n.endsWith(".docx"))
          .map(n -> n.substring(0, n.length() - ".docx".length())).sorted().toList();
    }
  }

  @ParameterizedTest(name = "{0}")
  @SuppressWarnings("null")
  @MethodSource("names")
  @DisplayName("lines have the same text at the same position as in Word's PDF")
  void sameLinesAsWord(String name) throws Exception {
    URL word = getClass().getResource("/ground-truth/" + name + ".docx");
    URL expectedPdf = getClass().getResource("/ground-truth/" + name + ".docx.pdf");
    Assumptions.assumeTrue(word != null && expectedPdf != null, "ground truth files exist");
    for (String font : FONTS) {
      Assumptions.assumeTrue(SystemFontLocator.findFontFile(font).isPresent(),
          "font is installed: " + font);
    }

    Path pdf = tempDir.resolve(name + ".pdf");
    WordToPdfUtil.generate(Path.of(word.toURI()), pdf,
        PdfGenerateOptions.builderForSystemFonts().build());

    List<List<Line>> expected = lines(Path.of(expectedPdf.toURI()));
    List<List<Line>> actual = lines(pdf);
    assertThat(actual).as("number of pages").hasSameSizeAs(expected);
    List<String> differences = new ArrayList<>();
    for (int p = 0; p < expected.size(); p++) {
      List<Line> e = expected.get(p);
      List<Line> a = actual.get(p);
      for (int i = 0; i < Math.max(e.size(), a.size()); i++) {
        String where = "page " + (p + 1) + " line " + (i + 1) + ": ";
        if (i >= e.size() || i >= a.size()) {
          differences.add(where + "Word has " + e.size() + " lines but " + a.size());
          break;
        }
        Line el = e.get(i);
        Line al = a.get(i);
        if (!el.text().equals(al.text())) {
          differences.add(where + "[" + al.text() + "] but Word has [" + el.text() + "]");
        } else if (Math.abs(el.baseline() - al.baseline()) > TOLERANCE
            || Math.abs(el.left() - al.left()) > TOLERANCE) {
          differences.add(where + "[" + al.text() + "] is at (" + al.left() + ", "
              + al.baseline() + ") but Word has it at (" + el.left() + ", " + el.baseline()
              + ")");
        }
      }
    }
    assertThat(differences).as("differences from Word's PDF").isEmpty();
  }

  /** Returns the lines of each page with their texts normalized. */
  private static List<List<Line>> lines(Path pdf) throws IOException {
    List<List<Line>> pages = new ArrayList<>();
    for (List<PdfLines.Line> page : PdfLines.of(pdf)) {
      List<Line> lines = new ArrayList<>();
      for (PdfLines.Line line : page) {
        lines.add(new Line(line.baseline(), line.left(), normalize(line.text())));
      }
      pages.add(lines);
    }
    return pages;
  }

  /**
   * Removes the differences which do not come from the layout: Word's PDF files use Kangxi
   * radicals and compatibility characters for some kanji, and the number of the characters of
   * a tab leader is not compared.
   */
  private static String normalize(String text) {
    return Normalizer.normalize(text, Normalizer.Form.NFKC).replace(" ", "")
        .replace('⺠', '民').replace('⻑', '長')
        .replace('ʻ', '‘').replace('ʼ', '’').replaceAll("[._-]{5,}", "…");
  }
}
