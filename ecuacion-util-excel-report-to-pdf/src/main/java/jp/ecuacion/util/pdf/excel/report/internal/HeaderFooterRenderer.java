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

import java.awt.Color;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.jspecify.annotations.Nullable;

/** Renders Excel headers and footers onto PDF pages. */
class HeaderFooterRenderer {

  /** A single formatted text run within a header or footer section. */
  private record HfRun(String text, boolean bold, float fontSize, Color color, boolean underline,
      boolean doubleUnderline, boolean strikethrough, boolean superscript, boolean subscript) {
  }

  private final FontManager fontManager;
  private final @Nullable Path sourcePath;

  HeaderFooterRenderer(FontManager fontManager, @Nullable Path sourcePath) {
    this.fontManager = fontManager;
    this.sourcePath = sourcePath;
  }

  /**
   * Renders the header or footer for the given page.
   *
   * @param cs the content stream to render into
   * @param sheet the sheet whose header/footer is rendered
   * @param pageSize the page rectangle
   * @param leftMargin left margin in points
   * @param rightMargin right margin in points
   * @param isHeader {@code true} to render the header, {@code false} for the footer
   * @param marginPt header or footer margin in points (distance from the page edge)
   * @param pageNumber 1-based current page number
   * @param totalPages total number of pages
   */
  void renderHeaderOrFooter(PDPageContentStream cs, Sheet sheet, PDRectangle pageSize,
      float leftMargin, float rightMargin, boolean isHeader, float marginPt, int pageNumber,
      int totalPages) throws IOException {

    String leftText = isHeader ? sheet.getHeader().getLeft() : sheet.getFooter().getLeft();
    String centerText = isHeader ? sheet.getHeader().getCenter() : sheet.getFooter().getCenter();
    String rightText = isHeader ? sheet.getHeader().getRight() : sheet.getFooter().getRight();

    if (isHfBlank(leftText) && isHfBlank(centerText) && isHfBlank(rightText)) {
      return;
    }

    String fileName = (sourcePath != null) ? hfFileNameWithoutExt(sourcePath) : "";
    String filePath =
        (sourcePath != null && sourcePath.getParent() != null) ? sourcePath.getParent().toString()
            : "";
    String sheetName = sheet.getSheetName();

    float defaultFontSize = 10f;
    PDType0Font defaultFont = fontManager.getFont(false);
    float ascent = defaultFont.getFontDescriptor().getAscent() / 1000f * defaultFontSize;
    float descent = defaultFont.getFontDescriptor().getDescent() / 1000f * defaultFontSize;

    float baseline =
        isHeader ? pageSize.getHeight() - marginPt - ascent : marginPt + Math.abs(descent);

    float pageWidth = pageSize.getWidth();

    if (!isHfBlank(leftText)) {
      List<HfRun> runs =
          parseHfRuns(leftText, pageNumber, totalPages, fileName, filePath, sheetName);
      renderHfSection(cs, runs, leftMargin, baseline);
    }
    if (!isHfBlank(centerText)) {
      List<HfRun> runs =
          parseHfRuns(centerText, pageNumber, totalPages, fileName, filePath, sheetName);
      float sectionWidth = computeHfRunsWidth(runs);
      renderHfSection(cs, runs, (pageWidth - sectionWidth) / 2f, baseline);
    }
    if (!isHfBlank(rightText)) {
      List<HfRun> runs =
          parseHfRuns(rightText, pageNumber, totalPages, fileName, filePath, sheetName);
      float sectionWidth = computeHfRunsWidth(runs);
      renderHfSection(cs, runs, pageWidth - rightMargin - sectionWidth, baseline);
    }
  }

  private void renderHfSection(PDPageContentStream cs, List<HfRun> runs, float startX,
      float baseline) throws IOException {
    float currentX = startX;
    for (HfRun run : runs) {
      if (run.text().isEmpty()) {
        continue;
      }
      float fs = (run.superscript() || run.subscript()) ? run.fontSize() * 0.7f : run.fontSize();
      float bl = baseline;
      if (run.superscript()) {
        bl += fs * 0.5f;
      } else if (run.subscript()) {
        bl -= fs * 0.2f;
      }
      PDType0Font font = fontManager.getFont(run.bold());
      float textWidth;
      try {
        textWidth = font.getStringWidth(run.text()) / 1000f * fs;
      } catch (Exception e) {
        currentX += fs;
        continue;
      }
      cs.beginText();
      cs.setFont(font, fs);
      cs.setNonStrokingColor(run.color());
      cs.newLineAtOffset(currentX, bl);
      try {
        cs.showText(run.text());
      } catch (Exception e) {
        // skip unrenderable text
      }
      cs.endText();
      float ascent = font.getFontDescriptor().getAscent() / 1000f * fs;
      float descent = font.getFontDescriptor().getDescent() / 1000f * fs;
      if (run.underline()) {
        float lineY = bl + descent - 0.5f;
        cs.setStrokingColor(run.color());
        cs.setLineWidth(0.5f);
        cs.moveTo(currentX, lineY);
        cs.lineTo(currentX + textWidth, lineY);
        cs.stroke();
      }
      if (run.doubleUnderline()) {
        float lineY1 = bl + descent - 0.5f;
        cs.setStrokingColor(run.color());
        cs.setLineWidth(0.5f);
        cs.moveTo(currentX, lineY1);
        cs.lineTo(currentX + textWidth, lineY1);
        cs.stroke();
        cs.moveTo(currentX, lineY1 - 1.5f);
        cs.lineTo(currentX + textWidth, lineY1 - 1.5f);
        cs.stroke();
      }
      if (run.strikethrough()) {
        float lineY = bl + (ascent + descent) / 2f;
        cs.setStrokingColor(run.color());
        cs.setLineWidth(0.5f);
        cs.moveTo(currentX, lineY);
        cs.lineTo(currentX + textWidth, lineY);
        cs.stroke();
      }
      currentX += textWidth;
    }
  }

  /**
   * Parses a header/footer section string into a list of {@link HfRun}s.
   *
   * <p>Handles all standard Excel header/footer format codes: {@code &P}, {@code &N},
   * {@code &D}, {@code &T}, {@code &A}, {@code &F}, {@code &Z}, {@code &B}, {@code &I},
   * {@code &U}, {@code &E}, {@code &S}, {@code &X}, {@code &Y}, {@code &"Font,Style"},
   * {@code &nn} (font size), {@code &KRRGGBB} (color), and {@code &&}.</p>
   *
   * <p>Italic ({@code &I}) is parsed but ignored because the embedded font has no
   * italic face. Picture insertion ({@code &G}) is not supported.</p>
   */
  private List<HfRun> parseHfRuns(String sectionText, int pageNumber, int totalPages,
      String fileName, String filePath, String sheetName) {
    List<HfRun> runs = new ArrayList<>();
    if (sectionText == null || sectionText.isBlank()) {
      return runs;
    }
    boolean bold = false;
    float fontSize = 10f;
    Color color = Color.BLACK;
    boolean underline = false;
    boolean doubleUnderline = false;
    boolean strikethrough = false;
    boolean superscript = false;
    boolean subscript = false;
    StringBuilder buf = new StringBuilder();
    int i = 0;
    while (i < sectionText.length()) {
      char c = sectionText.charAt(i);
      if (c != '&' || i + 1 >= sectionText.length()) {
        buf.append(c);
        i++;
        continue;
      }
      char code = Character.toUpperCase(sectionText.charAt(i + 1));
      switch (code) {
        case '&' -> {
          buf.append('&');
          i += 2;
        }
        case 'P' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          int offset = 0;
          int advance = i + 2;
          if (advance < sectionText.length()) {
            char mod = sectionText.charAt(advance);
            if ((mod == '+' || mod == '-') && advance + 1 < sectionText.length()
                && Character.isDigit(sectionText.charAt(advance + 1))) {
              int numEnd = advance + 1;
              while (numEnd < sectionText.length()
                  && Character.isDigit(sectionText.charAt(numEnd))) {
                numEnd++;
              }
              int n = Integer.parseInt(sectionText.substring(advance + 1, numEnd));
              offset = (mod == '+') ? n : -n;
              advance = numEnd;
            }
          }
          buf.append(pageNumber + offset);
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i = advance;
        }
        case 'N' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(totalPages);
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'D' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(LocalDate.now(ZoneId.systemDefault())
              .format(DateTimeFormatter.ofPattern("yyyy/M/d")));
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'T' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(LocalTime.now(ZoneId.systemDefault())
              .format(DateTimeFormatter.ofPattern("H:mm")));
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'A' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(sheetName);
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'F' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(fileName);
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'Z' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(filePath);
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'B' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          bold = !bold;
          i += 2;
        }
        case 'I' -> i += 2; // italic: no italic face available, skip
        case 'U' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          underline = !underline;
          if (underline) {
            doubleUnderline = false;
          }
          i += 2;
        }
        case 'E' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          doubleUnderline = !doubleUnderline;
          if (doubleUnderline) {
            underline = false;
          }
          i += 2;
        }
        case 'S' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          strikethrough = !strikethrough;
          i += 2;
        }
        case 'X' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          superscript = !superscript;
          if (superscript) {
            subscript = false;
          }
          i += 2;
        }
        case 'Y' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          subscript = !subscript;
          if (subscript) {
            superscript = false;
          }
          i += 2;
        }
        case '"' -> {
          int closeQ = sectionText.indexOf('"', i + 2);
          if (closeQ > i + 1) {
            flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
                superscript, subscript);
            String spec = sectionText.substring(i + 2, closeQ);
            int comma = spec.indexOf(',');
            if (comma >= 0) {
              bold = spec.substring(comma + 1).trim().equalsIgnoreCase("bold");
            }
            i = closeQ + 1;
          } else {
            i += 2;
          }
        }
        case 'K' -> {
          if (i + 7 <= sectionText.length()) {
            String hex = sectionText.substring(i + 2, i + 8);
            boolean allHex = hex.chars().allMatch(ch -> (ch >= '0' && ch <= '9')
                || (ch >= 'A' && ch <= 'F') || (ch >= 'a' && ch <= 'f'));
            if (allHex) {
              flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline,
                  strikethrough, superscript, subscript);
              color = new Color(Integer.parseInt(hex.substring(0, 2), 16),
                  Integer.parseInt(hex.substring(2, 4), 16),
                  Integer.parseInt(hex.substring(4, 6), 16));
              i += 8;
            } else {
              // Not a valid color code; skip &K
              i += 2;
            }
          } else {
            i += 2;
          }
        }
        default -> {
          // Try to parse &nn as a font size
          if (Character.isDigit(code)) {
            int numStart = i + 1;
            int numEnd = numStart;
            while (numEnd < sectionText.length() && Character.isDigit(sectionText.charAt(numEnd))) {
              numEnd++;
            }
            flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
                superscript, subscript);
            try {
              fontSize = Float.parseFloat(sectionText.substring(numStart, numEnd));
            } catch (NumberFormatException e) {
              // keep current size
            }
            i = numEnd;
          } else {
            i += 2; // unknown code, skip &X
          }
        }
      }
    }
    flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
        superscript, subscript);
    return runs;
  }

  private void flushHfRun(List<HfRun> runs, StringBuilder buf, boolean bold, float fontSize,
      Color color, boolean underline, boolean doubleUnderline, boolean strikethrough,
      boolean superscript, boolean subscript) {
    if (buf.isEmpty()) {
      return;
    }
    runs.add(new HfRun(buf.toString(), bold, fontSize, color, underline, doubleUnderline,
        strikethrough, superscript, subscript));
    buf.setLength(0);
  }

  private float computeHfRunsWidth(List<HfRun> runs) {
    float total = 0f;
    for (HfRun run : runs) {
      float fs = (run.superscript() || run.subscript()) ? run.fontSize() * 0.7f : run.fontSize();
      PDType0Font font = fontManager.getFont(run.bold());
      try {
        total += font.getStringWidth(run.text()) / 1000f * fs;
      } catch (Exception e) {
        total += fs;
      }
    }
    return total;
  }

  private boolean isHfBlank(@Nullable String s) {
    return s == null || s.isBlank();
  }

  private String hfFileNameWithoutExt(Path path) {
    String name = path.getFileName().toString();
    int dot = name.lastIndexOf('.');
    return (dot > 0) ? name.substring(0, dot) : name;
  }
}
