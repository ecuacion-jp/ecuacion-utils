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
package jp.ecuacion.util.pdf.excel.report.internal;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.core.api.Assertions.within;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("FontManager")
class FontManagerTest {

  private static Path regularFont() throws URISyntaxException {
    return Path.of(FontManagerTest.class
        .getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
  }

  private static Path boldFont() throws URISyntaxException {
    return Path.of(FontManagerTest.class
        .getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf").toURI());
  }

  @Nested
  @DisplayName("コンストラクタ (Path)")
  class PathConstructor {

    @Test
    @DisplayName("boldFontPath=null → bold/regular が同一フォントインスタンス")
    void noBoldPathFallsBackToRegular() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        assertThat(fm.getFont(true)).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("boldFontPath 指定 → bold と regular が別フォントインスタンス")
    void withBoldPathReturnsDifferentFonts() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), boldFont());
        assertThat(fm.getFont(true)).isNotSameAs(fm.getFont(false));
      }
    }
  }

  @Nested
  @DisplayName("コンストラクタ (TrueTypeFont)")
  class TtfConstructor {

    @Test
    @DisplayName("TrueTypeFont + boldTtf=null → bold/regular が同一フォント")
    void fromTtfNoBold() throws Exception {
      var ttf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, ttf, null, null, null);
        assertThat(fm.getFont(false)).isNotNull();
        assertThat(fm.getFont(true)).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("fallbackRegularPath 指定 → 構築成功、フォント非 null")
    void fromTtfWithFallbackPath() throws Exception {
      var ttf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, ttf, null, boldFont(), null);
        assertThat(fm.getFont(false)).isNotNull();
      }
    }

    @Test
    @DisplayName("bold TrueTypeFont 指定 → bold と regular が別フォント")
    void fromTtfWithBoldTtf() throws Exception {
      var regularTtf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      var boldTtf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(boldFont(), "Noto Sans JP Bold"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularTtf, boldTtf, null, null);
        assertThat(fm.getFont(true)).isNotSameAs(fm.getFont(false));
      }
    }
  }

  @Nested
  @DisplayName("getTypoAscent / getTypoDescent")
  class TypoMetrics {

    @Test
    @DisplayName("NotoSansJP: ascent > 0、descent < 0")
    void ascentAndDescent() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        assertThat(fm.getTypoAscent()).isGreaterThan(0f);
        assertThat(fm.getTypoDescent()).isLessThan(0f);
      }
    }
  }

  @Nested
  @DisplayName("selectFont")
  class SelectFont {

    @Test
    @DisplayName("NotoSansJP に含まれる文字 → primary font を返す")
    void encodableCharReturnsPrimaryFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), boldFont());
        var font = fm.selectFont('A', false);
        assertThat(font).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("bold=true・NotoSansJP に含まれる文字 → bold font を返す")
    void encodableCharReturnsBoldFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), boldFont());
        var font = fm.selectFont('A', true);
        assertThat(font).isSameAs(fm.getFont(true));
      }
    }

    @Test
    @DisplayName("fallback なし・エンコード不能文字 → PdfGenerateException")
    void noFallbackUnencodableThrows() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        // U+E000 は Private Use Area — いかなる標準フォントにも収録されない
        assertThatThrownBy(() -> fm.selectFont(0xE000, false))
            .isInstanceOf(PdfGenerateException.class);
      }
    }
  }

  @Nested
  @DisplayName("segmentText")
  class SegmentText {

    @Test
    @DisplayName("空文字列 → 空リスト")
    void emptyStringReturnsEmptyList() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        assertThat(fm.segmentText("", false)).isEmpty();
      }
    }

    @Test
    @DisplayName("ASCII テキスト → 結合すると元テキストと一致する TextRun リスト")
    void asciiTextReturnsRuns() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        List<FontManager.TextRun> runs = fm.segmentText("Hello", false);
        assertThat(runs).isNotEmpty();
        String combined = runs.stream().map(FontManager.TextRun::text).reduce("", String::concat);
        assertThat(combined).isEqualTo("Hello");
      }
    }

    @Test
    @DisplayName("bold=true → 各 TextRun が bold font を使う")
    void boldTextRunsUseBoldFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), boldFont());
        List<FontManager.TextRun> runs = fm.segmentText("Hi", true);
        assertThat(runs).isNotEmpty();
        runs.forEach(run -> assertThat(run.font()).isSameAs(fm.getFont(true)));
      }
    }
  }

  @Nested
  @DisplayName("getStringWidthWithFallback")
  class GetStringWidthWithFallback {

    @Test
    @DisplayName("通常テキスト → 正の幅")
    void normalTextReturnsPositiveWidth() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        float width = fm.getStringWidthWithFallback("Hello", false, 12f);
        assertThat(width).isGreaterThan(0f);
      }
    }

    @Test
    @DisplayName("空文字列 → 0")
    void emptyStringReturnsZero() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        assertThat(fm.getStringWidthWithFallback("", false, 12f)).isEqualTo(0f);
      }
    }

    @Test
    @DisplayName("エンコード不能文字 → fontSize を 1em として幅推定")
    void unencodableCharEstimatedAsFontSize() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularFont(), null);
        // U+E000 はエンコード不能 → 幅推定値 = fontSize (12f)
        String pua = new String(Character.toChars(0xE000));
        float width = fm.getStringWidthWithFallback(pua, false, 12f);
        assertThat(width).isCloseTo(12f, within(0.01f));
      }
    }
  }
}
