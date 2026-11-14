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

package jp.ecuacion.util.pdf.common.internal;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.util.pdf.common.exception.CharacterNotRenderableException;
import jp.ecuacion.util.pdf.common.exception.PdfGenerateException;
import org.apache.fontbox.ttf.HorizontalHeaderTable;
import org.apache.fontbox.ttf.NamingTable;
import org.apache.fontbox.ttf.OS2WindowsMetricsTable;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.jspecify.annotations.Nullable;

/**
 * Manages embedded fonts for PDF generation.
 *
 * <p>Loads fonts from the file system paths supplied via {@code
 * PdfGenerateOptions} of the module using this class.
 * If no bold font path is provided, the regular font is used for bold text as well.</p>
 *
 * <p>Regular and bold fonts are each configured as an ordered list. When a character cannot be
 * encoded by the primary font (the first entry), the rest of the list is tried, in order; for
 * bold text, the bold list is tried first, then the regular list. If the character is not
 * available in any configured font, {@link PdfGenerateException} is thrown.</p>
 *
 * <p>When {@link #enableSystemFontResolution(boolean)} is turned on, the font names used by
 * individual pieces of text (an Excel cell, a Word run; distinct from the document's default
 * font) can be resolved on demand via the {@code fontName}-accepting overloads below, so that
 * text using a different font than the document default (e.g. a CJK font for text written in
 * Japanese while the document's default font only covers Latin text) is rendered with its own,
 * more accurate font. Font names that cannot be resolved from the system fall back to the
 * document's default font.</p>
 */
@SuppressWarnings("exports")
public class FontManager {

  private static final DetailLogger detailLog = new DetailLogger(FontManager.class);

  /** A resolved regular/bold font pair, together with its metrics and diagnostic descriptions. */
  private record FontFamily(PDType0Font regular, PDType0Font bold, float typoAscent,
      float typoDescent, String regularDescription, String boldDescription) {
  }

  /** A loaded font together with a diagnostic description, used for fallback resolution. */
  private record LoadedFont(PDType0Font font, String description) {
  }

  /**
   * Vertical metrics of a font, as fractions of the em square, in the form Windows applications
   * derive their line heights from (the GDI {@code TEXTMETRIC} values).
   *
   * @param ascent          the height above the baseline (OS/2 {@code usWinAscent})
   * @param descent         the depth below the baseline, as a positive value (OS/2
   *                        {@code usWinDescent})
   * @param externalLeading the extra space between lines requested by the font, i.e. the part of
   *                        the {@code hhea} line gap not already covered by {@code ascent} and
   *                        {@code descent}
   * @param eastAsian       whether the font declares a CJK code page (OS/2
   *                        {@code ulCodePageRange1} bits 17-21)
   * @param hheaAscent      the ascender of the {@code hhea} table, which applications on macOS
   *                        derive their line heights from
   * @param hheaDescent     the descender of the {@code hhea} table, as a positive value
   * @param hheaLineGap     the line gap of the {@code hhea} table
   */
  public record LineMetrics(float ascent, float descent, float externalLeading,
      boolean eastAsian, float hheaAscent, float hheaDescent, float hheaLineGap) {
  }

  private static final LineMetrics FALLBACK_LINE_METRICS =
      new LineMetrics(0.8f, 0.2f, 0f, false, 0.8f, 0.2f, 0f);

  /** OS/2 ulCodePageRange1 bits 17-21: JIS/Japan, Simplified/Traditional Chinese, Korean. */
  private static final long EAST_ASIAN_CODE_PAGE_BITS = 0x003E0000L;

  private final PDDocument document;
  private final FontFamily defaultFamily;

  /** Line metrics of every font loaded by this instance. */
  private final IdentityHashMap<PDType0Font, LineMetrics> lineMetricsByFont =
      new IdentityHashMap<>();

  /** Regular fonts, tried in registration order (index 0 is the primary font). */
  private final List<LoadedFont> regularFonts;

  /** Bold fonts, tried in registration order, before falling through to {@link #regularFonts}. */
  private final List<LoadedFont> boldFonts;

  private boolean systemFontResolutionEnabled = false;

  private boolean exactBoldFaceRequired = false;

  /** Lazily-resolved, per-text font families, keyed by font name. Populated on first use. */
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
    registerLineMetrics(regularFonts, regularFontPaths);
    registerLineMetrics(boldFonts, boldFontPaths);
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
    lineMetricsByFont.put(regular, extractLineMetricsFromTtf(regularTtf));
    if (boldTtf != null) {
      lineMetricsByFont.put(bold, extractLineMetricsFromTtf(boldTtf));
    }
    this.regularFonts = loadFonts(document, fallbackRegularPaths);
    this.boldFonts = loadFonts(document, fallbackBoldPaths);
    registerLineMetrics(regularFonts, fallbackRegularPaths);
    registerLineMetrics(boldFonts, fallbackBoldPaths);
  }

  private void registerLineMetrics(List<LoadedFont> fonts, List<Path> fontPaths) {
    for (int i = 0; i < fonts.size(); i++) {
      lineMetricsByFont.put(fonts.get(i).font(), extractLineMetrics(fontPaths.get(i)));
    }
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
   * Enables or disables on-demand resolution of per-text font names via the system font
   * directories (see {@link SystemFontLocator}). When disabled (the default), every
   * {@code fontName}-accepting method behaves exactly like its no-argument counterpart, i.e.
   * always uses the document's default font.
   *
   * @param enabled {@code true} to resolve distinct per-text font names from the system
   */
  public void enableSystemFontResolution(boolean enabled) {
    this.systemFontResolutionEnabled = enabled;
  }

  /**
   * Sets whether the bold face of a font resolved from the system must belong to exactly the
   * same family. When {@code false} (the default), a bold face of a family with a similar
   * name can be taken.
   *
   * @param required {@code true} to take only the bold face of exactly the same family
   */
  public void requireExactBoldFace(boolean required) {
    this.exactBoldFaceRequired = required;
  }

  /**
   * Returns the font family for {@code fontName}, resolving and caching it from the system
   * font directories on first use. Falls back to the document's default font family (with a
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
   * bold-suffix lookup convention used for the document's default font. Returns the document's
   * default font family (with a warning logged) when the font cannot be found or loaded.
   */
  private FontFamily resolveNamedFamily(String fontName) {
    try {
      var fontFile = SystemFontLocator.findFontFile(fontName);
      if (fontFile.isEmpty()) {
        detailLog.warn("Font '" + fontName + "' used by the document was not found on the system; "
            + "falling back to the document's default font for the text using it.");
        return defaultFamily;
      }
      TrueTypeFont regularTtf = SystemFontLocator.loadTrueTypeFont(fontFile.get(), fontName);
      if (regularTtf == null) {
        detailLog.warn("Failed to load font '" + fontName + "' from " + fontFile.get()
            + "; falling back to the document's default font for the text using it.");
        return defaultFamily;
      }
      var boldFontFile = exactBoldFaceRequired
          ? SystemFontLocator.findExactFontFile(fontName + " Bold")
          : SystemFontLocator.findFontFile(fontName + " Bold");
      TrueTypeFont boldTtf = boldFontFile.isPresent()
          ? SystemFontLocator.loadTrueTypeFont(boldFontFile.get(), fontName + " Bold") : null;
      String regularDescription = describeFontOrFallback(regularTtf);
      String boldDescription =
          (boldTtf != null) ? describeFontOrFallback(boldTtf) : regularDescription;
      PDType0Font regular = PDType0Font.load(document, regularTtf, true);
      PDType0Font bold = (boldTtf != null) ? PDType0Font.load(document, boldTtf, true) : regular;
      float[] metrics = extractTypoMetricsFromTtf(regularTtf);
      lineMetricsByFont.put(regular, extractLineMetricsFromTtf(regularTtf));
      if (boldTtf != null) {
        lineMetricsByFont.put(bold, extractLineMetricsFromTtf(boldTtf));
      }
      return new FontFamily(regular, bold, metrics[0], metrics[1], regularDescription,
          boldDescription);
    } catch (IOException e) {
      detailLog.warn("Failed to resolve font '" + fontName
          + "' used by the document from the system: " + e.getMessage()
          + "; falling back to the document's default font for the text using it.");
      return defaultFamily;
    }
  }

  private static PDType0Font loadFontFromPath(PDDocument document, Path fontPath)
      throws IOException {
    if (fontPath.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".ttc")) {
      // A TrueType Collection: its first font is used.
      TrueTypeFont ttf = SystemFontLocator.loadTrueTypeFont(fontPath, "");
      if (ttf == null) {
        throw new IOException("No font is found in the font collection: " + fontPath);
      }
      return PDType0Font.load(document, ttf, true);
    }
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

  private static LineMetrics extractLineMetrics(Path fontPath) {
    try {
      TrueTypeFont ttf = SystemFontLocator.loadTrueTypeFont(fontPath, "");
      if (ttf == null) {
        return FALLBACK_LINE_METRICS;
      }
      try {
        return extractLineMetricsFromTtf(ttf);
      } finally {
        ttf.close();
      }
    } catch (IOException ignored) { // NOPMD
      return FALLBACK_LINE_METRICS;
    }
  }

  private static LineMetrics extractLineMetricsFromTtf(TrueTypeFont ttf) {
    try {
      float em = ttf.getUnitsPerEm();
      HorizontalHeaderTable hhea = ttf.getHorizontalHeader();
      if (em <= 0 || hhea == null) {
        return FALLBACK_LINE_METRICS;
      }
      float hheaHeight = (hhea.getAscender() - hhea.getDescender()) / em;
      float ascent = hhea.getAscender() / em;
      float descent = -hhea.getDescender() / em;
      boolean eastAsian = false;
      OS2WindowsMetricsTable os2 = ttf.getOS2Windows();
      if (os2 != null) {
        if (os2.getWinAscent() + os2.getWinDescent() > 0) {
          ascent = os2.getWinAscent() / em;
          descent = os2.getWinDescent() / em;
        }
        eastAsian = (os2.getCodePageRange1() & EAST_ASIAN_CODE_PAGE_BITS) != 0;
      }
      // Same derivation as GDI's tmExternalLeading: the hhea line gap, minus the part of it
      // the (usually larger) win ascent/descent already cover.
      float externalLeading =
          Math.max(0f, hhea.getLineGap() / em - ((ascent + descent) - hheaHeight));
      return new LineMetrics(ascent, descent, externalLeading, eastAsian,
          hhea.getAscender() / em, -hhea.getDescender() / em,
          Math.max(0f, hhea.getLineGap() / em));
    } catch (IOException ignored) { // NOPMD
      return FALLBACK_LINE_METRICS;
    }
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
      StringBuilder sb =
          new StringBuilder(family != null ? family : Objects.requireNonNull(postscript));
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
   * document's default font.
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
   * @param fontName the font family name of the text (e.g. an Excel cell's or a Word run's font)
   * @param bold     {@code true} for bold text, {@code false} for regular text
   * @return the PDF font
   */
  public PDType0Font getFont(String fontName, boolean bold) {
    FontFamily family = getFamily(fontName);
    return bold ? family.bold() : family.regular();
  }

  /**
   * Returns the appropriate font for the given Unicode code point, applying fallback logic,
   * from the document's default font.
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
   * @param fontName  the font family name of the text (e.g. an Excel cell's or a Word run's font)
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
   * {@link #selectFont(int, boolean)}, using the document's default font.
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
   * @param fontName the font family name of the text (e.g. an Excel cell's or a Word run's font)
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
   * selection to account for fallback fonts, from the document's default font.
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
   * @param fontName the font family name of the text (e.g. an Excel cell's or a Word run's font)
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
   * Returns the line metrics of a font obtained from this instance (e.g. via
   * {@link #selectFont(String, int, boolean)}).
   *
   * @param font a font returned by this instance
   * @return the font's line metrics; a generic estimate if they could not be read
   */
  public LineMetrics getLineMetrics(PDType0Font font) {
    return lineMetricsByFont.getOrDefault(font, FALLBACK_LINE_METRICS);
  }

  /**
   * Returns the typographic ascent in 1/1000 em units (from TTF OS/2 sTypoAscender), from the
   * document's default font. Use this for text positioning to match Excel's rendering.
   */
  public float getTypoAscent() {
    return defaultFamily.typoAscent();
  }

  /**
   * Same as {@link #getTypoAscent()}, but resolving {@code fontName} on demand when system font
   * resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the font family name of the text (e.g. an Excel cell's or a Word run's font)
   */
  public float getTypoAscent(String fontName) {
    return getFamily(fontName).typoAscent();
  }

  /**
   * Returns the typographic descent in 1/1000 em units (from TTF OS/2 sTypoDescender,
   * typically negative), from the document's default font. Use this for text positioning to
   * match Excel's rendering.
   */
  public float getTypoDescent() {
    return defaultFamily.typoDescent();
  }

  /**
   * Same as {@link #getTypoDescent()}, but resolving {@code fontName} on demand when system
   * font resolution is enabled (see {@link #enableSystemFontResolution(boolean)}).
   *
   * @param fontName the font family name of the text (e.g. an Excel cell's or a Word run's font)
   */
  public float getTypoDescent(String fontName) {
    return getFamily(fontName).typoDescent();
  }
}
