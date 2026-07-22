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
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.util.pdf.excel.report.exception.CharacterNotRenderableException;
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
 * <p>Regular and bold fonts are each configured as an ordered list. When a character cannot be
 * encoded by the primary font (the first entry), the rest of the list is tried, in order; for
 * bold text, the bold list is tried first, then the regular list. If the character is not
 * available in any configured font, {@link PdfGenerateException} is thrown.</p>
 *
 * <p>When {@link #enableSystemFontResolution(boolean)} is turned on, per-cell font names
 * (distinct from the workbook's default font) can be resolved on demand via the
 * {@code fontName}-accepting overloads below, so that cells using a different font than the
 * workbook default (e.g. a CJK font for cells written in Japanese while the workbook's default
 * font only covers Latin text) are rendered with their own, more accurate font. Font names that
 * cannot be resolved from the system fall back to the workbook's default font.</p>
 */
public class FontManager {

  private static final DetailLogger detailLog = new DetailLogger(FontManager.class);

  /** A resolved regular/bold font pair, together with its metrics and diagnostic descriptions. */
  private record FontFamily(PDType0Font regular, PDType0Font bold, float typoAscent,
      float typoDescent, String regularDescription, String boldDescription) {
  }

  /** A loaded font together with a diagnostic description, used for fallback resolution. */
  private record LoadedFont(PDType0Font font, String description) {
  }

  private final PDDocument document;
  private final FontFamily defaultFamily;

  /** Regular fonts, tried in registration order (index 0 is the primary font). */
  private final List<LoadedFont> regularFonts;

  /** Bold fonts, tried in registration order, before falling through to {@link #regularFonts}. */
  private final List<LoadedFont> boldFonts;

  private boolean systemFontResolutionEnabled = false;

  /** Lazily-resolved, per-cell font families, keyed by font name. Populated on first use. */
  private final Map<String, FontFamily> namedFamilyCache = new HashMap<>();

  /**
   * Constructs a {@code FontManager} loading fonts from file paths.
   *
   * @param document         the PDF document to embed fonts into
   * @param regularFontPaths paths to the TTF/TTC files used for regular text, tried in order
   *                         (must not be empty; the first entry is the primary font)
   * @param boldFontPaths    paths to the TTF/TTC files used for bold text, tried in order
   *                         before falling through to {@code regularFontPaths}. When empty,
   *                         {@code regularFontPaths} is used for bold text as well
   * @throws IOException if a font file cannot be read
   */
  public FontManager(PDDocument document, List<Path> regularFontPaths, List<Path> boldFontPaths)
      throws IOException {
    this.document = document;
    this.regularFonts = loadFonts(document, regularFontPaths);
    this.boldFonts = loadFonts(document, boldFontPaths);
    PDType0Font regular = regularFonts.get(0).font();
    PDType0Font bold = !boldFonts.isEmpty() ? boldFonts.get(0).font() : regular;
    String regularDescription = regularFonts.get(0).description();
    String boldDescription = !boldFonts.isEmpty() ? boldFonts.get(0).description()
        : regularDescription;
    float[] metrics = extractTypoMetrics(regularFontPaths.get(0));
    this.defaultFamily =
        new FontFamily(regular, bold, metrics[0], metrics[1], regularDescription, boldDescription);
  }

  /**
   * Constructs a {@code FontManager} loading its primary fonts from {@link TrueTypeFont} objects
   * (e.g. fonts resolved from a TrueType Collection via {@link SystemFontLocator}), with
   * additional fallback fonts loaded from file paths.
   *
   * @param document           the PDF document to embed fonts into
   * @param regularTtf         regular-weight primary font
   * @param boldTtf            bold-weight primary font, or {@code null} to fall back to
   *                           {@code regularTtf}
   * @param fallbackRegularPaths paths to fallback TTF/TTC files used for regular text, tried
   *                             in order after {@code regularTtf}
   * @param fallbackBoldPaths    paths to fallback TTF/TTC files used for bold text, tried in
   *                             order after {@code boldTtf}, before falling through to
   *                             {@code fallbackRegularPaths}
   * @throws IOException if a font cannot be loaded
   */
  public FontManager(PDDocument document, TrueTypeFont regularTtf, @Nullable TrueTypeFont boldTtf,
      List<Path> fallbackRegularPaths, List<Path> fallbackBoldPaths) throws IOException {
    this.document = document;
    // Extract naming info before the TrueTypeFont is consumed by PDType0Font.load.
    String regularDescription = describeFontOrFallback(regularTtf);
    String boldDescription =
        (boldTtf != null) ? describeFontOrFallback(boldTtf) : regularDescription;
    PDType0Font regular = PDType0Font.load(document, regularTtf, true);
    PDType0Font bold = (boldTtf != null) ? PDType0Font.load(document, boldTtf, true) : regular;
    float[] metrics = extractTypoMetricsFromTtf(regularTtf);
    this.defaultFamily =
        new FontFamily(regular, bold, metrics[0], metrics[1], regularDescription, boldDescription);
    this.regularFonts = loadFonts(document, fallbackRegularPaths);
    this.boldFonts = loadFonts(document, fallbackBoldPaths);
  }

  private static List<LoadedFont> loadFonts(PDDocument document, List<Path> fontPaths)
      throws IOException {
    List<LoadedFont> result = new ArrayList<>();
    for (Path fontPath : fontPaths) {
      result.add(new LoadedFont(loadFontFromPath(document, fontPath), describeFontFile(fontPath)));
    }
    return result;
  }

  /**
   * Enables or disables on-demand resolution of per-cell font names via the system font
   * directories (see {@link SystemFontLocator}). When disabled (the default), every
   * {@code fontName}-accepting method behaves exactly like its no-argument counterpart, i.e.
   * always uses the workbook's default font.
   *
   * @param enabled {@code true} to resolve distinct per-cell font names from the system
   */
  public void enableSystemFontResolution(boolean enabled) {
    this.systemFontResolutionEnabled = enabled;
  }

  /**
   * Returns the font family for {@code fontName}, resolving and caching it from the system
   * font directories on first use. Falls back to the workbook's default font family (with a
   * diagnostic warning) when {@code fontName} cannot be resolved, or when system font
   * resolution is disabled.
   */
  private FontFamily getFamily(String fontName) {
    if (!systemFontResolutionEnabled) {
      return defaultFamily;
    }
    FontFamily cached = namedFamilyCache.get(fontName);
    if (cached != null) {
      return cached;
    }
    FontFamily resolved = resolveNamedFamily(fontName);
    namedFamilyCache.put(fontName, resolved);
    return resolved;
  }

  /**
   * Resolves {@code fontName} from the system font directories, following the same TTC /
   * bold-suffix lookup convention used for the workbook's default font. Returns the workbook's
   * default font family (with a warning logged) when the font cannot be found or loaded.
   */
  private FontFamily resolveNamedFamily(String fontName) {
    try {
      var fontFile = SystemFontLocator.findFontFile(fontName);
      if (fontFile.isEmpty()) {
        detailLog.warn("Font '" + fontName + "' used by a cell was not found on the system; "
            + "falling back to the workbook's default font for this cell.");
        return defaultFamily;
      }
      TrueTypeFont regularTtf = SystemFontLocator.loadTrueTypeFont(fontFile.get(), fontName);
      if (regularTtf == null) {
        detailLog.warn("Failed to load font '" + fontName + "' from " + fontFile.get()
            + "; falling back to the workbook's default font for this cell.");
        return defaultFamily;
      }
      var boldFontFile = SystemFontLocator.findFontFile(fontName + " Bold");
      TrueTypeFont boldTtf = boldFontFile.isPresent()
          ? SystemFontLocator.loadTrueTypeFont(boldFontFile.get(), fontName + " Bold") : null;
      String regularDescription = describeFontOrFallback(regularTtf);
      String boldDescription =
          (boldTtf != null) ? describeFontOrFallback(boldTtf) : regularDescription;
      PDType0Font regular = PDType0Font.load(document, regularTtf, true);
      PDType0Font bold = (boldTtf != null) ? PDType0Font.load(document, boldTtf, true) : regular;
      float[] metrics = extractTypoMetricsFromTtf(regularTtf);
      return new FontFamily(regular, bold, metrics[0], metrics[1], regularDescription,
          boldDescription);
    } catch (IOException e) {
      detailLog.warn("Failed to resolve font '" + fontName + "' used by a cell from the system: "
          + e.getMessage() + "; falling back to the workbook's default font for this cell.");
      return defaultFamily;
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
   * Returns the font to use for the given boldness (primary font, no fallback), from the
   * workbook's default font.
   *
   * @param bold {@code true} for bold text, {@code false} for regular text
   * @return the PDF font
   */
  public PDType0Font getFont(boolean bold) {
    return bold ? defaultFamily.bold() : defaultFamily.regular();
  }

  /**
   * Returns the font to use for the given boldness, resolving {@code fontName} on demand when
   * system font resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the cell's font family name (e.g. from {@code Font.getFontName()})
   * @param bold     {@code true} for bold text, {@code false} for regular text
   * @return the PDF font
   */
  public PDType0Font getFont(String fontName, boolean bold) {
    FontFamily family = getFamily(fontName);
    return bold ? family.bold() : family.regular();
  }

  /**
   * Returns the appropriate font for the given Unicode code point, applying fallback logic,
   * from the workbook's default font.
   *
   * <p>If the primary font can encode the character, it is returned. Otherwise the fallback
   * fonts are tried, in order. If none can encode the character, {@link PdfGenerateException} is
   * thrown so the caller can surface an actionable error.</p>
   *
   * @param codePoint the Unicode code point to render
   * @param bold      {@code true} for bold weight
   * @return the font that can encode {@code codePoint}
   * @throws PdfGenerateException if no configured font can encode the character
   */
  public PDType0Font selectFont(int codePoint, boolean bold) {
    return selectFontFromFamily(defaultFamily, codePoint, bold);
  }

  /**
   * Same as {@link #selectFont(int, boolean)}, but resolving {@code fontName} on demand when
   * system font resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName  the cell's font family name (e.g. from {@code Font.getFontName()})
   * @param codePoint the Unicode code point to render
   * @param bold      {@code true} for bold weight
   * @return the font that can encode {@code codePoint}
   * @throws PdfGenerateException if no configured font can encode the character
   */
  public PDType0Font selectFont(String fontName, int codePoint, boolean bold) {
    return selectFontFromFamily(getFamily(fontName), codePoint, bold);
  }

  private PDType0Font selectFontFromFamily(FontFamily family, int codePoint, boolean bold) {
    PDType0Font primary = bold ? family.bold() : family.regular();
    if (canEncode(primary, codePoint)) {
      return primary;
    }
    if (bold) {
      for (LoadedFont font : boldFonts) {
        if (canEncode(font.font(), codePoint)) {
          return font.font();
        }
      }
    }
    for (LoadedFont font : regularFonts) {
      if (canEncode(font.font(), codePoint)) {
        return font.font();
      }
    }
    StringBuilder fallbackDescription = new StringBuilder();
    if (bold && !boldFonts.isEmpty()) {
      fallbackDescription.append("bold: ").append(
          boldFonts.stream().map(LoadedFont::description).collect(Collectors.joining(", ")));
    }
    if (!regularFonts.isEmpty()) {
      if (fallbackDescription.length() > 0) {
        fallbackDescription.append("; ");
      }
      fallbackDescription.append("regular: ").append(
          regularFonts.stream().map(LoadedFont::description).collect(Collectors.joining(", ")));
    }
    if (fallbackDescription.length() == 0) {
      fallbackDescription.append("(none configured)");
    }
    String primaryDescription = bold ? family.boldDescription() : family.regularDescription();
    throw new CharacterNotRenderableException(
        Integer.toHexString(codePoint).toUpperCase(Locale.ROOT),
        new String(Character.toChars(codePoint)), primaryDescription,
        fallbackDescription.toString());
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
   * {@link #selectFont(int, boolean)}, using the workbook's default font.
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
  public List<TextRun> segmentText(String text, boolean bold) {
    return segmentTextFromFamily(defaultFamily, text, bold);
  }

  /**
   * Same as {@link #segmentText(String, boolean)}, but resolving {@code fontName} on demand
   * when system font resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the cell's font family name (e.g. from {@code Font.getFontName()})
   * @param text     the string to segment
   * @param bold     {@code true} for bold weight
   * @return ordered list of text runs, each with its assigned font
   * @throws PdfGenerateException if a character cannot be rendered by any configured font
   */
  public List<TextRun> segmentText(String fontName, String text, boolean bold) {
    return segmentTextFromFamily(getFamily(fontName), text, bold);
  }

  private List<TextRun> segmentTextFromFamily(FontFamily family, String text, boolean bold) {
    List<TextRun> runs = new ArrayList<>();
    if (text.isEmpty()) {
      return runs;
    }
    StringBuilder current = new StringBuilder();
    PDType0Font currentFont = null;

    for (int i = 0; i < text.length();) {
      int cp = text.codePointAt(i);
      PDType0Font font = selectFontFromFamily(family, cp, bold);
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
   * selection to account for fallback fonts, from the workbook's default font.
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
    return getStringWidthWithFallbackFromFamily(defaultFamily, text, bold, fontSize);
  }

  /**
   * Same as {@link #getStringWidthWithFallback(String, boolean, float)}, but resolving
   * {@code fontName} on demand when system font resolution is enabled (see
   * {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the cell's font family name (e.g. from {@code Font.getFontName()})
   * @param text     the string whose width to measure
   * @param bold     {@code true} for bold weight
   * @param fontSize the font size in points
   * @return total advance width in points
   */
  public float getStringWidthWithFallback(String fontName, String text, boolean bold,
      float fontSize) {
    return getStringWidthWithFallbackFromFamily(getFamily(fontName), text, bold, fontSize);
  }

  private float getStringWidthWithFallbackFromFamily(FontFamily family, String text, boolean bold,
      float fontSize) {
    float total = 0f;
    for (int i = 0; i < text.length();) {
      int cp = text.codePointAt(i);
      PDType0Font font;
      try {
        font = selectFontFromFamily(family, cp, bold);
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
   * Returns the typographic ascent in 1/1000 em units (from TTF OS/2 sTypoAscender), from the
   * workbook's default font. Use this for text positioning to match Excel's rendering.
   */
  public float getTypoAscent() {
    return defaultFamily.typoAscent();
  }

  /**
   * Same as {@link #getTypoAscent()}, but resolving {@code fontName} on demand when system font
   * resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the cell's font family name (e.g. from {@code Font.getFontName()})
   */
  public float getTypoAscent(String fontName) {
    return getFamily(fontName).typoAscent();
  }

  /**
   * Returns the typographic descent in 1/1000 em units (from TTF OS/2 sTypoDescender,
   * typically negative), from the workbook's default font. Use this for text positioning to
   * match Excel's rendering.
   */
  public float getTypoDescent() {
    return defaultFamily.typoDescent();
  }

  /**
   * Same as {@link #getTypoDescent()}, but resolving {@code fontName} on demand when system
   * font resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the cell's font family name (e.g. from {@code Font.getFontName()})
   */
  public float getTypoDescent(String fontName) {
    return getFamily(fontName).typoDescent();
  }
}
