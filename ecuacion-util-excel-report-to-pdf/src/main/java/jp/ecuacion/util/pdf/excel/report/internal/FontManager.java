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

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import org.apache.fontbox.ttf.NamingTable;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.jspecify.annotations.Nullable;

/**
 * Manages embedded fonts for PDF generation.
 *
 * <p>Loads fonts from the file system paths supplied via {@link
 * jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions}.
 * If no bold font path is provided, the regular font is used for bold text as well.</p>
 *
 * <p>An optional fallback font pair can be configured. When a character cannot be encoded
 * by the primary font, the fallback font is tried. If the character is not available in
 * either font, {@link PdfGenerateException} is thrown.</p>
 */
public class FontManager {

  private final PDType0Font regularFont;
  private final PDType0Font boldFont;
  @Nullable
  private final PDType0Font fallbackRegularFont;
  @Nullable
  private final PDType0Font fallbackBoldFont;

  /** Human-readable identification of each loaded font, for diagnostics in error messages. */
  private final String regularFontDescription;

  private final String boldFontDescription;
  @Nullable
  private final String fallbackRegularFontDescription;
  @Nullable
  private final String fallbackBoldFontDescription;

  /**
   * Typographic ascent in 1/1000 em units (from TTF OS/2 sTypoAscender).
   *
   * <p>Excel positions text using sTypo metrics rather than the larger usWinAscent values
   * stored in PDFBox's font descriptor. Using sTypo metrics produces line spacing that
   * matches Excel's rendering, especially for CJK fonts whose usWinAscent can be &gt; 1em.</p>
   */
  private final float typoAscent;

  /** Typographic descent in 1/1000 em units (from TTF OS/2 sTypoDescender, typically negative). */
  private final float typoDescent;

  /**
   * Constructs a {@code FontManager} loading fonts from file paths (no fallback font).
   *
   * @param document        the PDF document to embed fonts into
   * @param regularFontPath path to the TTF/TTC file used for regular text
   * @param boldFontPath    path to the TTF/TTC file used for bold text, or {@code null} to
   *                        fall back to {@code regularFontPath}
   * @throws IOException if a font file cannot be read
   */
  public FontManager(PDDocument document, Path regularFontPath, @Nullable Path boldFontPath)
      throws IOException {
    regularFont = loadFontFromPath(document, regularFontPath);
    boldFont = (boldFontPath != null) ? loadFontFromPath(document, boldFontPath) : regularFont;
    fallbackRegularFont = null;
    fallbackBoldFont = null;
    regularFontDescription = describeFontFile(regularFontPath);
    boldFontDescription =
        (boldFontPath != null) ? describeFontFile(boldFontPath) : regularFontDescription;
    fallbackRegularFontDescription = null;
    fallbackBoldFontDescription = null;
    float[] metrics = extractTypoMetrics(regularFontPath);
    typoAscent = metrics[0];
    typoDescent = metrics[1];
  }

  /**
   * Constructs a {@code FontManager} loading fonts from {@link TrueTypeFont} objects
   * (e.g. fonts resolved from a TrueType Collection via {@link SystemFontLocator}),
   * with optional fallback fonts for characters not covered by the primary fonts.
   *
   * @param document            the PDF document to embed fonts into
   * @param regularTtf          regular-weight font
   * @param boldTtf             bold-weight font, or {@code null} to fall back to
   *                            {@code regularTtf}
   * @param fallbackRegularPath path to the fallback TTF/TTC used when the primary font
   *                            lacks a glyph, or {@code null} for no fallback
   * @param fallbackBoldPath    path to the bold fallback font, or {@code null} to share
   *                            {@code fallbackRegularPath}
   * @throws IOException if a font cannot be loaded
   */
  public FontManager(PDDocument document, TrueTypeFont regularTtf, @Nullable TrueTypeFont boldTtf,
      @Nullable Path fallbackRegularPath, @Nullable Path fallbackBoldPath) throws IOException {
    // Extract naming info before the TrueTypeFont is consumed by PDType0Font.load.
    regularFontDescription = describeFontOrFallback(regularTtf);
    boldFontDescription =
        (boldTtf != null) ? describeFontOrFallback(boldTtf) : regularFontDescription;
    regularFont = PDType0Font.load(document, regularTtf, true);
    boldFont = (boldTtf != null) ? PDType0Font.load(document, boldTtf, true) : regularFont;
    float[] metrics = extractTypoMetricsFromTtf(regularTtf);
    typoAscent = metrics[0];
    typoDescent = metrics[1];
    if (fallbackRegularPath != null) {
      fallbackRegularFont = loadFontFromPath(document, fallbackRegularPath);
      fallbackBoldFont = (fallbackBoldPath != null) ? loadFontFromPath(document, fallbackBoldPath)
          : fallbackRegularFont;
      fallbackRegularFontDescription = describeFontFile(fallbackRegularPath);
      fallbackBoldFontDescription = (fallbackBoldPath != null) ? describeFontFile(fallbackBoldPath)
          : fallbackRegularFontDescription;
    } else {
      fallbackRegularFont = null;
      fallbackBoldFont = null;
      fallbackRegularFontDescription = null;
      fallbackBoldFontDescription = null;
    }
  }

  private static PDType0Font loadFontFromPath(PDDocument document, Path fontPath)
      throws IOException {
    try (InputStream is = Files.newInputStream(fontPath)) {
      return PDType0Font.load(document, is, true);
    }
  }

  private static float[] extractTypoMetricsFromTtf(TrueTypeFont ttf) {
    try {
      int em = ttf.getUnitsPerEm();
      if (em <= 0) {
        return fallbackMetrics();
      }
      int ascender = ttf.getOS2Windows().getTypoAscender();
      int descender = ttf.getOS2Windows().getTypoDescender();
      if (ascender == 0 && descender == 0) {
        return fallbackMetrics();
      }
      return new float[] {(float) ascender / em * 1000, (float) descender / em * 1000};
    } catch (IOException ignored) { // NOPMD
      return fallbackMetrics();
    }
  }

  private static float[] extractTypoMetrics(Path fontPath) {
    try {
      TrueTypeFont ttf = SystemFontLocator.loadTrueTypeFont(fontPath, "");
      if (ttf == null) {
        return fallbackMetrics();
      }
      float[] result = extractTypoMetricsFromTtf(ttf);
      ttf.close();
      return result;
    } catch (IOException ignored) { // NOPMD
      return fallbackMetrics();
    }
  }

  private static float[] fallbackMetrics() {
    return new float[] {800f, -200f};
  }

  /**
   * Loads {@code fontPath} and builds a diagnostic description from its name table, falling
   * back to the bare path when the file cannot be parsed or has no usable name records.
   */
  private static String describeFontFile(Path fontPath) {
    try {
      TrueTypeFont ttf = SystemFontLocator.loadTrueTypeFont(fontPath, "");
      if (ttf == null) {
        return fontPath.toString();
      }
      try {
        String naming = describeFont(ttf);
        return (naming != null) ? naming + " [" + fontPath + "]" : fontPath.toString();
      } finally {
        ttf.close();
      }
    } catch (IOException ignored) { // NOPMD
      return fontPath.toString();
    }
  }

  /**
   * Same as {@link #describeFont(TrueTypeFont)}, but returns a generic placeholder instead of
   * {@code null} when the font has no usable name records (there is no file path to fall back
   * to, unlike {@link #describeFontFile(Path)}).
   */
  private static String describeFontOrFallback(TrueTypeFont ttf) {
    String description = describeFont(ttf);
    return (description != null) ? description : "(unknown font)";
  }

  /**
   * Builds a human-readable description (family, subfamily/style, PostScript name) from a
   * TrueType/OpenType font's {@code name} table, for identifying which font was actually
   * resolved when a glyph cannot be found (e.g. a font matched by name but lacking CJK coverage).
   *
   * @return a description string, or {@code null} if the font has no usable name records
   */
  @Nullable
  private static String describeFont(TrueTypeFont ttf) {
    try {
      NamingTable naming = ttf.getNaming();
      if (naming == null) {
        return null;
      }
      String family = null;
      String subfamily = null;
      String postscript = null;
      for (var record : naming.getNameRecords()) {
        String value = record.getString();
        if (value == null || value.isBlank()) {
          continue;
        }
        switch (record.getNameId()) {
          case 1 -> family = (family == null) ? value : family;
          case 2 -> subfamily = (subfamily == null) ? value : subfamily;
          case 6 -> postscript = (postscript == null) ? value : postscript;
          default -> {
          }
        }
      }
      if (family == null && postscript == null) {
        return null;
      }
      StringBuilder sb = new StringBuilder(family != null ? family : postscript);
      if (subfamily != null) {
        sb.append(' ').append(subfamily);
      }
      if (postscript != null) {
        sb.append(" (PostScript: ").append(postscript).append(')');
      }
      return sb.toString();
    } catch (IOException ignored) { // NOPMD
      return null;
    }
  }

  /**
   * Returns the font to use for the given boldness (primary font, no fallback).
   *
   * @param bold {@code true} for bold text, {@code false} for regular text
   * @return the PDF font
   */
  public PDType0Font getFont(boolean bold) {
    return bold ? boldFont : regularFont;
  }

  /**
   * Returns the appropriate font for the given Unicode code point, applying fallback logic.
   *
   * <p>If the primary font can encode the character, it is returned. Otherwise the fallback
   * font is tried. If neither can encode the character, {@link PdfGenerateException} is
   * thrown so the caller can surface an actionable error.</p>
   *
   * @param codePoint the Unicode code point to render
   * @param bold      {@code true} for bold weight
   * @return the font that can encode {@code codePoint}
   * @throws PdfGenerateException if no configured font can encode the character
   */
  public PDType0Font selectFont(int codePoint, boolean bold) throws PdfGenerateException {
    PDType0Font primary = bold ? boldFont : regularFont;
    if (canEncode(primary, codePoint)) {
      return primary;
    }
    PDType0Font fallback = bold ? fallbackBoldFont : fallbackRegularFont;
    if (fallback != null && canEncode(fallback, codePoint)) {
      return fallback;
    }
    String primaryDescription = bold ? boldFontDescription : regularFontDescription;
    String fallbackDescription =
        bold ? fallbackBoldFontDescription : fallbackRegularFontDescription;
    throw new PdfGenerateException(
        "Character U+" + Integer.toHexString(codePoint).toUpperCase(Locale.ROOT) + " ('"
            + new String(Character.toChars(codePoint)) + "') cannot be rendered: "
            + "glyph not available in any configured font. " + "Primary font used: "
            + primaryDescription + ". "
            + (fallbackDescription != null ? "Fallback font used: " + fallbackDescription + "."
                : "No fallback font is configured.")
            + " Add a font that covers this character via PdfGenerateOptions.");
  }

  private static boolean canEncode(PDType0Font font, int codePoint) {
    try {
      font.getStringWidth(new String(Character.toChars(codePoint)));
      return true;
    } catch (Exception ignored) { // NOPMD
      return false;
    }
  }

  /**
   * A contiguous run of text that should be rendered with the same font.
   *
   * @param font the font to use for this run
   * @param text the text content of this run
   */
  public record TextRun(PDType0Font font, String text) {
  }

  /**
   * Splits {@code text} into runs, each assigned the appropriate font via
   * {@link #selectFont(int, boolean)}.
   *
   * <p>Consecutive characters that map to the same font are grouped into a single run.
   * Throws {@link PdfGenerateException} if any character cannot be encoded by either
   * the primary or fallback font.</p>
   *
   * @param text the string to segment
   * @param bold {@code true} for bold weight
   * @return ordered list of text runs, each with its assigned font
   * @throws PdfGenerateException if a character cannot be rendered by any configured font
   */
  public List<TextRun> segmentText(String text, boolean bold) throws PdfGenerateException {
    List<TextRun> runs = new ArrayList<>();
    if (text.isEmpty()) {
      return runs;
    }
    StringBuilder current = new StringBuilder();
    PDType0Font currentFont = null;

    for (int i = 0; i < text.length();) {
      int cp = text.codePointAt(i);
      PDType0Font font = selectFont(cp, bold);
      if (currentFont == null) {
        currentFont = font;
      }
      if (!font.equals(currentFont)) {
        runs.add(new TextRun(currentFont, current.toString()));
        current.setLength(0);
        currentFont = font;
      }
      current.appendCodePoint(cp);
      i += Character.charCount(cp);
    }
    if (currentFont != null && current.length() > 0) {
      runs.add(new TextRun(currentFont, current.toString()));
    }
    return runs;
  }

  /**
   * Computes the total advance width of {@code text} in points, using per-character font
   * selection to account for fallback fonts.
   *
   * <p>If a character cannot be encoded, a best-effort estimate of {@code fontSize} is used
   * for its width rather than throwing, so layout decisions remain stable even when individual
   * characters are unrenderable.</p>
   *
   * @param text     the string whose width to measure
   * @param bold     {@code true} for bold weight
   * @param fontSize the font size in points
   * @return total advance width in points
   */
  public float getStringWidthWithFallback(String text, boolean bold, float fontSize) {
    float total = 0f;
    for (int i = 0; i < text.length();) {
      int cp = text.codePointAt(i);
      PDType0Font font;
      try {
        font = selectFont(cp, bold);
      } catch (PdfGenerateException e) {
        total += fontSize; // unrenderable character: estimate one em
        i += Character.charCount(cp);
        continue;
      }
      try {
        total += font.getStringWidth(new String(Character.toChars(cp))) / 1000f * fontSize;
      } catch (Exception e) {
        total += fontSize;
      }
      i += Character.charCount(cp);
    }
    return total;
  }

  /**
   * Returns the typographic ascent in 1/1000 em units (from TTF OS/2 sTypoAscender).
   * Use this for text positioning to match Excel's rendering.
   */
  public float getTypoAscent() {
    return typoAscent;
  }

  /**
   * Returns the typographic descent in 1/1000 em units (from TTF OS/2 sTypoDescender,
   * typically negative). Use this for text positioning to match Excel's rendering.
   */
  public float getTypoDescent() {
    return typoDescent;
  }
}
