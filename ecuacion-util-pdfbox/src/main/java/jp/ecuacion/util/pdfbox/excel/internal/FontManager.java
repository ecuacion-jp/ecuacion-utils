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
package jp.ecuacion.util.pdfbox.excel.internal;

import java.io.IOException;
import java.io.InputStream;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.font.PDType0Font;

/**
 * Manages embedded fonts for PDF generation.
 *
 * <p>Loads Noto Sans JP Regular and Bold from the classpath.
 * Font files must be placed under {@code src/main/resources/fonts/NotoSansJP/}:
 * <ul>
 *   <li>{@code NotoSansJP-Regular.ttf} — used for regular text</li>
 *   <li>{@code NotoSansJP-Bold.ttf} — used for bold text</li>
 * </ul>
 * </p>
 */
public class FontManager {

  private static final String FONT_REGULAR = "/fonts/NotoSansJP/NotoSansJP-Regular.ttf";
  private static final String FONT_BOLD = "/fonts/NotoSansJP/NotoSansJP-Bold.ttf";

  private final PDType0Font regularFont;
  private final PDType0Font boldFont;

  /**
   * Constructs a {@code FontManager} and loads fonts into the given document.
   *
   * @param document the PDF document to embed fonts into
   * @throws IOException if a font resource cannot be found or loaded
   */
  public FontManager(PDDocument document) throws IOException {
    regularFont = loadFont(document, FONT_REGULAR);
    boldFont = loadFont(document, FONT_BOLD);
  }

  private PDType0Font loadFont(PDDocument document, String resourcePath) throws IOException {
    try (InputStream is = getClass().getResourceAsStream(resourcePath)) {
      if (is == null) {
        throw new IOException("Font resource not found: " + resourcePath
            + ". Place the TTF file under src/main/resources" + resourcePath);
      }
      return PDType0Font.load(document, is, true);
    }
  }

  /**
   * Returns the font to use for the given boldness.
   *
   * @param bold {@code true} for bold (IPAex Mincho), {@code false} for regular (IPAex Gothic)
   * @return the PDF font
   */
  public PDType0Font getFont(boolean bold) {
    return bold ? boldFont : regularFont;
  }
}
