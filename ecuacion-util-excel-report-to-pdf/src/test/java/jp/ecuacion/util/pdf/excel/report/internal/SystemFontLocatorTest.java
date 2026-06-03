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
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Path;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for {@link SystemFontLocator}. */
@DisplayName("SystemFontLocator")
class SystemFontLocatorTest {

  // NotoSansJP-Regular: '0' glyph advance=555, unitsPerEm=1000, at 11pt@96dpi → 8.14px
  // round(8.14)=8 (correct), ceil(8.14)=9 (wrong for NotoSansJP)
  private static final float FONT_SIZE_PT = 11.0f;

  private static Path notoSansJpRegularPath() throws URISyntaxException {
    var url = SystemFontLocatorTest.class.getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf");
    return Path.of(url.toURI());
  }

  @Nested
  @DisplayName("computeMdw")
  class ComputeMdw {

    @Test
    @DisplayName("uses Math.round(), not Math.ceil() — NotoSansJP '0' advance=8.14px → MDW=8")
    void usesRoundNotCeil() throws URISyntaxException {
      Path fontPath = notoSansJpRegularPath();

      int mdw = SystemFontLocator.computeMdw(fontPath, "Noto Sans JP", FONT_SIZE_PT);

      // round(8.14) = 8; ceil(8.14) = 9 — verify round() is used, not ceil()
      assertThat(mdw).isEqualTo(8);
    }

    @Test
    @DisplayName("returns a positive value for a valid font file")
    void returnsPositive() throws URISyntaxException {
      Path fontPath = notoSansJpRegularPath();

      int mdw = SystemFontLocator.computeMdw(fontPath, "Noto Sans JP", FONT_SIZE_PT);

      assertThat(mdw).isGreaterThan(0);
    }

    @Test
    @DisplayName("returns positive fallback value when font file does not exist")
    void returnsFallbackForMissingFile(@TempDir Path tempDir) {
      Path nonExistent = tempDir.resolve("no_such_font.ttf");

      int mdw = SystemFontLocator.computeMdw(nonExistent, "FakeFont", FONT_SIZE_PT);

      assertThat(mdw).isGreaterThan(0);
    }
  }

  @Nested
  @DisplayName("computeExcelMdw")
  class ComputeExcelMdw {

    @Test
    @DisplayName("uses Math.ceil(), not Math.round() — NotoSansJP '0' advance=8.14px → MDW=9")
    void usesCeilNotRound() throws URISyntaxException {
      Path fontPath = notoSansJpRegularPath();

      int mdw = SystemFontLocator.computeExcelMdw(fontPath, "Noto Sans JP", FONT_SIZE_PT);

      // ceil(8.14) = 9; round(8.14) = 8 — verify ceil() is used (matches Excel pixel measurement)
      assertThat(mdw).isEqualTo(9);
    }

    @Test
    @DisplayName("returns a positive value for a valid font file")
    void returnsPositive() throws URISyntaxException {
      Path fontPath = notoSansJpRegularPath();

      int mdw = SystemFontLocator.computeExcelMdw(fontPath, "Noto Sans JP", FONT_SIZE_PT);

      assertThat(mdw).isGreaterThan(0);
    }
  }

  @Nested
  @DisplayName("loadTrueTypeFont")
  class LoadTrueTypeFont {

    @Test
    @DisplayName("loads a TTF file and returns a non-null font with valid unitsPerEm")
    void loadsTtfFile() throws URISyntaxException, IOException {
      Path fontPath = notoSansJpRegularPath();

      var ttf = SystemFontLocator.loadTrueTypeFont(fontPath, "Noto Sans JP");

      assertThat(ttf).isNotNull();
      if (ttf != null) {
        assertThat(ttf.getUnitsPerEm()).isGreaterThan(0);
      }
    }

    @Test
    @DisplayName("loads correct subfont from TTC when PostScript name differs (Meiryo UI case)")
    void loadsTtcByFamilyName() throws IOException {
      // Simulate a TTC where getFontByName(familyName) returns null because the PostScript name
      // differs from the family name (e.g. "Meiryo UI" family but "MeiryoUI" PostScript name).
      // We test this by creating a TTC copy and verifying that the correct subfont is selected.
      // If Meiryo is present, verify that "Meiryo UI" loads the proportional variant (~58% em).
      var meiryo = SystemFontLocator.findFontFile("Meiryo UI");
      org.junit.jupiter.api.Assumptions.assumeTrue(meiryo.isPresent(),
          "Meiryo UI not installed — skipping TTC subfamily test");

      var ttf = SystemFontLocator.loadTrueTypeFont(meiryo.get(), "Meiryo UI");

      assertThat(ttf).isNotNull();
      if (ttf != null) {
        // Meiryo UI has proportional kana: 'ア' (U+30A1) advance ≈ 58% of em (≤ 70%)
        // Meiryo (full-width) has 'ア' at 100% of em.
        var cmap = ttf.getUnicodeCmapLookup();
        int gid = cmap != null ? cmap.getGlyphId(0x30A1) : 0;
        int advance = ttf.getHorizontalMetrics().getAdvanceWidth(gid);
        int unitsPerEm = ttf.getUnitsPerEm();
        double ratio = (double) advance / unitsPerEm;
        assertThat(ratio).as("Meiryo UI 'ア' should be proportional (<70% em), not full-width")
            .isLessThan(0.70);
      }
    }
  }

  @Nested
  @DisplayName("findFontFile")
  class FindFontFile {

    @Test
    @DisplayName("returns empty for a font name that is not installed on this system")
    void returnsEmptyForUnknownFont() {
      var result = SystemFontLocator.findFontFile("__NonExistentFontXyz123__");

      assertThat(result).isEmpty();
    }

    @Test
    @DisplayName("游ゴシック search returns Medium over Light and Regular (Word/Excel install)")
    void prefersYuGothicMediumOverLight() {
      var result = SystemFontLocator.findFontFile("游ゴシック");
      org.junit.jupiter.api.Assumptions.assumeTrue(result.isPresent(),
          "游ゴシック not found — skipping weight-preference test");

      // findFontFile must return the Medium variant (score=0), not Regular (score=1) or
      // Light (score=3).  Excel ships only YuGothR (Regular) in its DFonts, while Word
      // ships YuGothM (Medium) as well.  The cross-directory ranking must pick Medium
      // over Regular so that our PDF matches the heavier weight Excel uses via CoreText.
      int score = SystemFontLocator.getRegularStyleScore(result.get(), "游ゴシック");
      assertThat(score)
          .as("游ゴシック should resolve to Medium (score=0), not Regular (1) or Light (3)")
          .isEqualTo(0);
    }
  }

  @Nested
  @DisplayName("getRegularStyleScore")
  class GetRegularStyleScore {

    @Test
    @DisplayName("NotoSansJP Regular → score 1")
    void regularFontReturnsScore1() throws URISyntaxException {
      Path fontPath = notoSansJpRegularPath();

      int score = SystemFontLocator.getRegularStyleScore(fontPath, "Noto Sans JP");

      assertThat(score).isEqualTo(1);
    }

    @Test
    @DisplayName("NotoSansJP Bold → score 10 (bold heavily penalised)")
    void boldFontReturnsScore10() throws URISyntaxException {
      var url = SystemFontLocatorTest.class.getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf");
      Path fontPath = Path.of(url.toURI());

      int score = SystemFontLocator.getRegularStyleScore(fontPath, "Noto Sans JP Bold");

      assertThat(score).isEqualTo(10);
    }
  }
}
