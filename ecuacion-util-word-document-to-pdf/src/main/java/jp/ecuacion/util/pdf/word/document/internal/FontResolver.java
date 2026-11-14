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
package jp.ecuacion.util.pdf.word.document.internal;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import jp.ecuacion.util.pdf.common.internal.FontManager;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.jspecify.annotations.Nullable;

/**
 * Chooses the font for each character of a run, following the rules of Word: a run has
 * separate fonts for ASCII, East Asian and the other characters.
 */
final class FontResolver {

  /**
   * Holds a font chosen for a character.
   *
   * @param font the font
   * @param syntheticBold whether bold is required but the font is not a bold face
   * @param ascent the height above the baseline used for line heights, in em
   * @param descent the depth below the baseline used for line heights, in em
   */
  record Face(PDType0Font font, boolean syntheticBold, float ascent, float descent) {
  }

  private final FontManager fontManager;
  private final String defaultFontName;
  private final Map<String, Map<Integer, Face>> faces = new HashMap<>();
  private final IdentityHashMap<PDType0Font, Map<Integer, Float>> widths =
      new IdentityHashMap<>();

  FontResolver(FontManager fontManager, String defaultFontName) {
    this.fontManager = fontManager;
    this.defaultFontName = defaultFontName;
  }

  /**
   * Returns the font for the character.
   *
   * @throws jp.ecuacion.util.pdf.common.exception.CharacterNotRenderableException if no font
   *     has the glyph
   */
  Face face(RunProps rp, int cp) {
    Map<Integer, Face> byChar =
        Objects.requireNonNull(faces.computeIfAbsent(rp.fontKey(), k -> new HashMap<>()));
    return Objects.requireNonNull(byChar.computeIfAbsent(cp, k -> resolve(rp, cp)));
  }

  private Face resolve(RunProps rp, int cp) {
    final boolean bold = rp.isBold();
    List<String> names = new ArrayList<>();
    add(names, slotFont(rp, cp));
    add(names, rp.eastAsia);
    add(names, rp.ascii);
    add(names, rp.hansi);
    if (names.isEmpty()) {
      names.add(defaultFontName);
    }
    for (String name : names) {
      PDType0Font font = fontManager.getFont(name, bold);
      if (!Float.isNaN(width(font, cp))) {
        return toFace(font, bold && font.equals(fontManager.getFont(name, false)));
      }
    }
    // Not in the fonts of the run: the fallback fonts are tried (an exception if none has it).
    PDType0Font font = fontManager.selectFont(names.get(0), cp, bold);
    return toFace(font, bold && font.equals(fontManager.selectFont(names.get(0), cp, false)));
  }

  private static void add(List<String> names, @Nullable String name) {
    if (name != null && !names.contains(name)) {
      names.add(name);
    }
  }

  private static @Nullable String slotFont(RunProps rp, int cp) {
    if (cp < 0x80) {
      return rp.ascii;
    }
    if (TextRules.eastAsian(cp)) {
      return rp.eastAsia;
    }
    // Characters used both in Western and East Asian text (quotes, symbols, Greek etc.)
    // use the East Asian font only when the run says so.
    return "eastAsia".equals(rp.hint) ? rp.eastAsia : rp.hansi;
  }

  private Face toFace(PDType0Font font, boolean syntheticBold) {
    // The metrics of the hhea table are used, which matches the output of Word for Mac
    // (verified against PDF files exported by it). Word for Windows uses the OS/2 win
    // metrics instead, which differ for some fonts (e.g. Yu Mincho).
    FontManager.LineMetrics lm = fontManager.getLineMetrics(font);
    float ascent = lm.hheaAscent();
    float descent = lm.hheaDescent();
    if (lm.eastAsian()) {
      // Word gives East Asian fonts 30% more line height than their ascent + descent.
      float extra = (ascent + descent) * 0.3f;
      ascent += extra / 2;
      descent += extra / 2;
    } else {
      ascent += lm.hheaLineGap();
    }
    return new Face(font, syntheticBold, ascent, descent);
  }

  /** Returns the advance width in em, or NaN if the font does not have the character. */
  float width(PDType0Font font, int cp) {
    Map<Integer, Float> map =
        Objects.requireNonNull(widths.computeIfAbsent(font, k -> new HashMap<>()));
    Float w = map.get(cp);
    if (w == null) {
      try {
        w = font.getStringWidth(new String(Character.toChars(cp))) / 1000f;
      } catch (Exception e) { // NOPMD - any failure means the font cannot show the character
        w = Float.NaN;
      }
      map.put(cp, w);
    }
    return w;
  }
}
