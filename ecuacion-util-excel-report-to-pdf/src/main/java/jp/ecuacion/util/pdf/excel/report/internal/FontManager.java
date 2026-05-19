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
 */
public class FontManager {

  private final PDType0Font regularFont;
  private final PDType0Font boldFont;

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
   * Constructs a {@code FontManager} loading fonts from file paths.
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
    float[] metrics = extractTypoMetrics(regularFontPath);
    typoAscent = metrics[0];
    typoDescent = metrics[1];
  }

  /**
   * Constructs a {@code FontManager} loading fonts from {@link TrueTypeFont} objects
   * (e.g. fonts resolved from a TrueType Collection via {@link SystemFontLocator}).
   *
   * @param document        the PDF document to embed fonts into
   * @param regularTtf      regular-weight font
   * @param boldTtf         bold-weight font, or {@code null} to fall back to {@code regularTtf}
   * @throws IOException if a font cannot be loaded
   */
  public FontManager(PDDocument document, TrueTypeFont regularTtf, @Nullable TrueTypeFont boldTtf)
      throws IOException {
    regularFont = PDType0Font.load(document, regularTtf, true);
    boldFont = (boldTtf != null) ? PDType0Font.load(document, boldTtf, true) : regularFont;
    float[] metrics = extractTypoMetricsFromTtf(regularTtf);
    typoAscent = metrics[0];
    typoDescent = metrics[1];
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
   * Returns the font to use for the given boldness.
   *
   * @param bold {@code true} for bold text, {@code false} for regular text
   * @return the PDF font
   */
  public PDType0Font getFont(boolean bold) {
    return bold ? boldFont : regularFont;
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
