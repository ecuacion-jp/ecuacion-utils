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
package jp.ecuacion.util.pdf.word.document.util;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.jspecify.annotations.Nullable;

/** Extracts the lines of a PDF file with their positions, for assertions on the layout. */
final class PdfLines {

  /**
   * A line of text. Positions are in points from the top-left of the page.
   *
   * @param baseline the y of the baseline
   * @param left the left edge of the first glyph
   * @param right the right edge of the last glyph
   * @param text the text without spaces
   * @param glyphs the glyphs from left to right
   */
  record Line(float baseline, float left, float right, String text, List<TextPosition> glyphs) {

    /** Returns the left edge of the first occurrence of the text in the line. */
    float leftOf(String part) {
      int index = text.indexOf(part);
      if (index < 0) {
        throw new AssertionError("[" + part + "] is not in [" + text + "]");
      }
      int length = 0;
      for (TextPosition g : glyphs) {
        if (length >= index) {
          return g.getXDirAdj();
        }
        length += g.getUnicode().length();
      }
      throw new AssertionError("[" + part + "] is not found in the glyphs");
    }

    /** Returns the right edge of the first occurrence of the text in the line. */
    float rightOf(String part) {
      int end = text.indexOf(part) + part.length();
      int length = 0;
      for (TextPosition g : glyphs) {
        length += g.getUnicode().length();
        if (length >= end) {
          return g.getXDirAdj() + g.getWidthDirAdj();
        }
      }
      throw new AssertionError("[" + part + "] is not found in the glyphs");
    }
  }

  private PdfLines() {}

  /** Returns the lines of each page, from top to bottom. */
  static List<List<Line>> of(Path pdf) throws IOException {
    List<List<Line>> pages = new ArrayList<>();
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      for (int page = 1; page <= doc.getNumberOfPages(); page++) {
        List<TextPosition> glyphs = new ArrayList<>();
        PDFTextStripper stripper = new PDFTextStripper() {
          @Override
          protected void processTextPosition(@Nullable TextPosition text) {
            if (text != null && !text.getUnicode().isBlank()) {
              glyphs.add(text);
            }
          }
        };
        stripper.setStartPage(page);
        stripper.setEndPage(page);
        stripper.getText(doc);
        pages.add(toLines(glyphs));
      }
    }
    return pages;
  }

  /** Returns the line of the first page which contains the text. */
  static Line find(List<List<Line>> pages, String part) {
    for (List<Line> page : pages) {
      for (Line line : page) {
        if (line.text().contains(part)) {
          return line;
        }
      }
    }
    throw new AssertionError("no line contains [" + part + "]");
  }

  /** Groups the glyphs on the same baseline into lines. */
  private static List<Line> toLines(List<TextPosition> glyphs) {
    glyphs.sort(Comparator.comparingDouble(TextPosition::getYDirAdj));
    List<List<TextPosition>> groups = new ArrayList<>();
    for (TextPosition g : glyphs) {
      List<TextPosition> last = groups.isEmpty() ? null : groups.get(groups.size() - 1);
      if (last != null && Math.abs(last.get(0).getYDirAdj() - g.getYDirAdj()) < 1.5f) {
        last.add(g);
      } else {
        List<TextPosition> group = new ArrayList<>();
        group.add(g);
        groups.add(group);
      }
    }
    List<Line> lines = new ArrayList<>();
    for (List<TextPosition> group : groups) {
      group.sort(Comparator.comparingDouble(TextPosition::getXDirAdj));
      StringBuilder sb = new StringBuilder();
      for (TextPosition g : group) {
        sb.append(g.getUnicode());
      }
      TextPosition first = group.get(0);
      TextPosition last = group.get(group.size() - 1);
      lines.add(new Line(first.getYDirAdj(), first.getXDirAdj(),
          last.getXDirAdj() + last.getWidthDirAdj(), sb.toString(), List.copyOf(group)));
    }
    return lines;
  }
}
