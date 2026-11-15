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

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.pdf.common.exception.FontLoadFailedException;
import jp.ecuacion.util.pdf.common.exception.SystemFontNotFoundException;
import jp.ecuacion.util.pdf.common.internal.FontManager;
import jp.ecuacion.util.pdf.common.internal.SystemFontLocator;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;

/** Renders a Word file to a PDF document. */
public final class WordRenderer {

  /** The font Word uses when a document specifies no font at all. */
  private static final String LAST_RESORT_FONT = "Times New Roman";

  private WordRenderer() {}

  /**
   * Renders the Word file and adds its pages to the PDF document.
   *
   * @param wordPath the path to the Word (.docx) file
   * @param document the PDF document to add pages to
   * @param options the options, which decide the fonts
   * @throws IOException if the Word file or a font file cannot be read
   */
  public static void render(Path wordPath, PDDocument document, PdfGenerateOptions options)
      throws IOException {
    Model.Document doc;
    try (DocxPackage pkg = DocxPackage.open(wordPath, options.getWordPassword())) {
      doc = DocxReader.read(pkg);
    }
    List<String> names = new ArrayList<>();
    for (String name : new String[] {doc.defaultRun.eastAsia, doc.defaultRun.ascii,
        doc.defaultRun.hansi, LAST_RESORT_FONT}) {
      if (name != null && !names.contains(name)) {
        names.add(name);
      }
    }
    FontManager fontManager = createFontManager(document, options, names);
    FontResolver fonts = new FontResolver(fontManager, names.get(0));
    Layout.PassResult result = new DocumentLayouter(doc, fonts).layout();
    new PdfPainter(document, result).paint();
  }

  private static FontManager createFontManager(PDDocument document, PdfGenerateOptions options,
      List<String> defaultFontNames) throws IOException {
    FontManager fontManager = null;
    if (options.isUseSystemFonts()) {
      for (String name : defaultFontNames) {
        var file = SystemFontLocator.findFontFile(name);
        if (file.isEmpty()) {
          continue;
        }
        TrueTypeFont regular = SystemFontLocator.loadTrueTypeFont(file.get(), name);
        if (regular == null) {
          throw new FontLoadFailedException(name, file.get());
        }
        // Like Word, only the bold face of the same family is used. Without it, bold text
        // is the regular face emboldened.
        var boldFile = SystemFontLocator.findExactFontFile(name + " Bold");
        TrueTypeFont bold = boldFile.isPresent()
            ? SystemFontLocator.loadTrueTypeFont(boldFile.get(), name + " Bold") : null;
        fontManager = new FontManager(document, regular, bold, options.getRegularFontPaths(),
            options.getBoldFontPaths());
        break;
      }
      if (fontManager == null && options.getRegularFontPaths().isEmpty()) {
        throw new SystemFontNotFoundException(defaultFontNames.get(0));
      }
    }
    if (fontManager == null) {
      fontManager = new FontManager(document, options.getRegularFontPaths(),
          options.getBoldFontPaths());
    }
    // Fonts of each run are looked up on the system. A font which is not found is replaced
    // by the default font above.
    fontManager.enableSystemFontResolution(options.isUseSystemFonts());
    fontManager.requireExactBoldFace(true);
    return fontManager;
  }
}
