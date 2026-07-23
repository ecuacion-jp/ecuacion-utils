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
import jp.ecuacion.util.pdf.excel.report.exception.CharacterNotRenderableException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.assertj.core.api.InstanceOfAssertFactories;
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
  @DisplayName("constructor (Path)")
  class PathConstructor {

    @Test
    @DisplayName("boldFontPaths empty — bold and regular are the same font instance")
    void noBoldPathFallsBackToRegular() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        assertThat(fm.getFont(true)).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("boldFontPaths set — bold and regular are different font instances")
    void withBoldPathReturnsDifferentFonts() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        assertThat(fm.getFont(true)).isNotSameAs(fm.getFont(false));
      }
    }
  }

  @Nested
  @DisplayName("constructor (TrueTypeFont)")
  class TtfConstructor {

    @Test
    @DisplayName("TrueTypeFont + boldTtf=null — bold and regular are the same font")
    void fromTtfNoBold() throws Exception {
      var ttf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, ttf, null, List.of(), List.of());
        assertThat(fm.getFont(false)).isNotNull();
        assertThat(fm.getFont(true)).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("fallbackRegularPaths set — construction succeeds, font non-null")
    void fromTtfWithFallbackPath() throws Exception {
      var ttf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, ttf, null, List.of(boldFont()), List.of());
        assertThat(fm.getFont(false)).isNotNull();
      }
    }

    @Test
    @DisplayName("bold TrueTypeFont set — bold and regular are different fonts")
    void fromTtfWithBoldTtf() throws Exception {
      var regularTtf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      var boldTtf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(boldFont(), "Noto Sans JP Bold"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, regularTtf, boldTtf, List.of(), List.of());
        assertThat(fm.getFont(true)).isNotSameAs(fm.getFont(false));
      }
    }
  }

  @Nested
  @DisplayName("getTypoAscent / getTypoDescent")
  class TypoMetrics {

    @Test
    @DisplayName("NotoSansJP: ascent > 0, descent < 0")
    void ascentAndDescent() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        assertThat(fm.getTypoAscent()).isGreaterThan(0f);
        assertThat(fm.getTypoDescent()).isLessThan(0f);
      }
    }
  }

  @Nested
  @DisplayName("selectFont")
  class SelectFont {

    @Test
    @DisplayName("encodable char in NotoSansJP — returns primary font")
    void encodableCharReturnsPrimaryFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        var font = fm.selectFont('A', false);
        assertThat(font).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("bold=true, encodable char in NotoSansJP — returns bold font")
    void encodableCharReturnsBoldFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        var font = fm.selectFont('A', true);
        assertThat(font).isSameAs(fm.getFont(true));
      }
    }

    @Test
    @DisplayName("no fallback, unencodable char — throws CharacterNotRenderableException")
    void noFallbackUnencodableThrows() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        // U+E000 is Private Use Area — not included in any standard font
        assertThatThrownBy(() -> fm.selectFont(0xE000, false))
            .isInstanceOf(CharacterNotRenderableException.class);
      }
    }
  }

  @Nested
  @DisplayName("segmentText")
  class SegmentText {

    @Test
    @DisplayName("empty string — returns empty list")
    void emptyStringReturnsEmptyList() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        assertThat(fm.segmentText("", false)).isEmpty();
      }
    }

    @Test
    @DisplayName("ASCII text — concatenated TextRuns match original string")
    void asciiTextReturnsRuns() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        List<FontManager.TextRun> runs = fm.segmentText("Hello", false);
        assertThat(runs).isNotEmpty();
        String combined = runs.stream().map(FontManager.TextRun::text).reduce("", String::concat);
        assertThat(combined).isEqualTo("Hello");
      }
    }

    @Test
    @DisplayName("bold=true — each TextRun uses bold font")
    void boldTextRunsUseBoldFont() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        List<FontManager.TextRun> runs = fm.segmentText("Hi", true);
        assertThat(runs).isNotEmpty();
        runs.forEach(run -> assertThat(run.font()).isSameAs(fm.getFont(true)));
      }
    }
  }

  @Nested
  @DisplayName("named font resolution (per-cell fonts)")
  class NamedFontResolution {

    @Test
    @DisplayName("resolution disabled (default) — fontName is ignored, same as no-arg overload")
    void disabledResolutionIgnoresFontName() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        assertThat(fm.getFont("Some Other Font", false)).isSameAs(fm.getFont(false));
        assertThat(fm.getFont("Some Other Font", true)).isSameAs(fm.getFont(true));
        assertThat(fm.getTypoAscent("Some Other Font")).isEqualTo(fm.getTypoAscent());
        assertThat(fm.getTypoDescent("Some Other Font")).isEqualTo(fm.getTypoDescent());
      }
    }

    @Test
    @DisplayName("resolution enabled, unresolvable font name — falls back to default font")
    void unknownFontNameFallsBackToDefault() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        fm.enableSystemFontResolution(true);
        String fictionalName = "__FictionalFontForTestZZZ999__";
        assertThat(fm.getFont(fictionalName, false)).isSameAs(fm.getFont(false));
        assertThat(fm.getFont(fictionalName, true)).isSameAs(fm.getFont(true));
      }
    }

    @Test
    @DisplayName("resolution enabled, unresolvable font name — typo metrics fall back to default")
    void unknownFontNameTypoMetricsFallBackToDefault() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        fm.enableSystemFontResolution(true);
        String fictionalName = "__FictionalFontForTestZZZ999__";
        assertThat(fm.getTypoAscent(fictionalName)).isEqualTo(fm.getTypoAscent());
        assertThat(fm.getTypoDescent(fictionalName)).isEqualTo(fm.getTypoDescent());
      }
    }

    @Test
    @DisplayName("resolution enabled, unresolvable font name — selectFont/segmentText still work")
    void unknownFontNameSelectAndSegmentFallBack() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of(boldFont()));
        fm.enableSystemFontResolution(true);
        String fictionalName = "__FictionalFontForTestZZZ999__";
        assertThat(fm.selectFont(fictionalName, 'A', false)).isSameAs(fm.getFont(false));
        List<FontManager.TextRun> runs = fm.segmentText(fictionalName, "Hi", false);
        assertThat(runs).isNotEmpty();
        runs.forEach(run -> assertThat(run.font()).isSameAs(fm.getFont(false)));
        assertThat(fm.getStringWidthWithFallback(fictionalName, "Hi", false, 12f))
            .isEqualTo(fm.getStringWidthWithFallback("Hi", false, 12f));
      }
    }
  }

  @Nested
  @DisplayName("multiple fallback fonts")
  class MultipleFallbackFonts {

    @Test
    @DisplayName("Path constructor with multiple regular/bold entries — construction succeeds")
    void multipleFallbacksConstructSuccessfully() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc,
            List.of(regularFont(), regularFont(), regularFont()),
            List.of(boldFont(), boldFont()));
        assertThat(fm.getFont(false)).isNotNull();
      }
    }

    @Test
    @DisplayName("TrueTypeFont constructor with multiple fallback entries — construction succeeds")
    void multipleFallbacksFromTtfConstructor() throws Exception {
      var ttf = Objects.requireNonNull(
          SystemFontLocator.loadTrueTypeFont(regularFont(), "Noto Sans JP"));
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, ttf, null,
            List.of(regularFont(), regularFont()), List.of(boldFont()));
        assertThat(fm.getFont(false)).isNotNull();
      }
    }

    @Test
    @DisplayName("encodable char — resolved by primary font, fallbacks not needed")
    void encodableCharUsesPrimary() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc,
            List.of(regularFont(), regularFont()), List.of(boldFont()));
        assertThat(fm.selectFont('A', false)).isSameAs(fm.getFont(false));
      }
    }

    @Test
    @DisplayName("regular: all fallbacks fail — throws with regular fallback descriptions listed")
    void allRegularFallbacksFailThrowsWithDescriptions() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc,
            List.of(regularFont(), regularFont(), regularFont()), List.of());
        // U+E000 is Private Use Area — not included in any standard font, including
        // NotoSansJP, so neither the primary nor any of the fallback entries can encode it.
        assertThatThrownBy(() -> fm.selectFont(0xE000, false))
            .asInstanceOf(InstanceOfAssertFactories.throwable(CharacterNotRenderableException.class))
            .satisfies(e -> assertThat(e.getViolations().getBusinessViolations().get(0)
                .getMessageArgs()[3]).asString().contains("regular:"));
      }
    }

    @Test
    @DisplayName("bold: bold fallbacks exhausted — message lists both bold and regular fonts "
        + "tried (bold falls through to the regular list)")
    void allBoldFallbacksFailFallsThroughToRegularDescriptions() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc,
            List.of(regularFont(), regularFont()), List.of(boldFont(), boldFont()));
        // U+E000 is unencodable by every configured font (regular and bold alike), so the
        // exception's fallback description must show that both the bold list and the regular
        // fallback list (consulted after the bold list is exhausted) were tried.
        assertThatThrownBy(() -> fm.selectFont(0xE000, true))
            .asInstanceOf(InstanceOfAssertFactories.throwable(CharacterNotRenderableException.class))
            .satisfies(e -> assertThat(e.getViolations().getBusinessViolations().get(0)
                .getMessageArgs()[3]).asString().contains("bold:").contains("regular:"));
      }
    }
  }

  @Nested
  @DisplayName("getStringWidthWithFallback")
  class GetStringWidthWithFallback {

    @Test
    @DisplayName("normal text — returns positive width")
    void normalTextReturnsPositiveWidth() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        float width = fm.getStringWidthWithFallback("Hello", false, 12f);
        assertThat(width).isGreaterThan(0f);
      }
    }

    @Test
    @DisplayName("empty string — returns 0")
    void emptyStringReturnsZero() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        assertThat(fm.getStringWidthWithFallback("", false, 12f)).isEqualTo(0f);
      }
    }

    @Test
    @DisplayName("unencodable char — width estimated as fontSize (1em)")
    void unencodableCharEstimatedAsFontSize() throws Exception {
      try (PDDocument doc = new PDDocument()) {
        var fm = new FontManager(doc, List.of(regularFont()), List.of());
        // U+E000 is unencodable — estimated width = fontSize (12f)
        String pua = new String(Character.toChars(0xE000));
        float width = fm.getStringWidthWithFallback(pua, false, 12f);
        assertThat(width).isCloseTo(12f, within(0.01f));
      }
    }
  }
}
