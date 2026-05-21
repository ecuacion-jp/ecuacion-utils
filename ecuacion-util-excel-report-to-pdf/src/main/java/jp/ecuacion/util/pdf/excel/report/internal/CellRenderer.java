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
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.util.Matrix;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DifferentialStyleProvider;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.TableStyle;
import org.apache.poi.ss.usermodel.TableStyleType;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STHorizontalAlignment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STVerticalAlignment;

/** Renders Excel cells (background, text, borders) and resolves table styles. */
class CellRenderer {

  private static final float CELL_PADDING = 2f;
  private static final float INDENT_WIDTH_PT = 7f * (72f / 96f);

  private final FontManager fontManager;
  private final CellValueFormatter formatter;
  /** Set at the start of each render pass; used for theme colour resolution. */
  @Nullable XSSFWorkbook currentWorkbook;

  CellRenderer(FontManager fontManager, CellValueFormatter formatter) {
    this.fontManager = fontManager;
    this.formatter = formatter;
  }

  /** Renders background fill only (pass 1 of the two-pass row rendering). */
  void renderBackground(PDPageContentStream cs, @Nullable Cell cell, float x, float y,
      float width, float height, @Nullable TableCellStyle tableStyle) throws IOException {

    if (tableStyle != null && tableStyle.fill() != null) {
      cs.setNonStrokingColor(tableStyle.fill());
      cs.addRect(x, y, width, height);
      cs.fill();
    }

    CellStyle style = (cell != null) ? cell.getCellStyle() : null;
    if (style != null) {
      Color bgColor = getBackgroundColor(style);
      if (bgColor != null) {
        cs.setNonStrokingColor(bgColor);
        cs.addRect(x, y, width, height);
        cs.fill();
      }
    }
  }

  /**
   * Renders text and borders only (pass 2 of the two-pass row rendering).
   *
   * @param overflowWidth maximum horizontal width the text may occupy, including adjacent empty
   *        cells. Equal to {@code width} when no overflow is allowed.
   */
  void renderForeground(PDPageContentStream cs, @Nullable Cell cell, float x, float y,
      float width, float height, float scaleFactor, @Nullable TableCellStyle tableStyle,
      float overflowWidth) throws IOException, PdfGenerateException {

    if (cell != null) {
      String value = formatter.getCellDisplayValue(cell);
      if (value != null && !value.isBlank()) {
        if (cell.getCellStyle().getRotation() == 255) {
          renderVerticalText(cs, cell, value, x, y, width, height, scaleFactor);
        } else {
          Color tableFontColor = (tableStyle != null) ? tableStyle.fontColor() : null;
          boolean tableFontBold = (tableStyle != null) && tableStyle.fontBold();
          if (overflowWidth > width) {
            // Text may overflow into adjacent empty cells: clip to overflowWidth, not cellWidth.
            cs.saveGraphicsState();
            cs.addRect(x, y - 2, overflowWidth, height + 4); // +4 for descent/ascent safety
            cs.clip();
            renderText(cs, cell, value, x, y, overflowWidth, height, scaleFactor,
                tableFontColor, tableFontBold);
            cs.restoreGraphicsState();
          } else {
            renderText(cs, cell, value, x, y, width, height, scaleFactor,
                tableFontColor, tableFontBold);
          }
        }
      }
    }

    CellStyle style = (cell != null) ? cell.getCellStyle() : null;
    if (style != null && cell != null) {
      renderBorders(cs, style, cell.getSheet(), x, y, width, height);
    }

    if (tableStyle != null) {
      renderTableBorderLine(cs, tableStyle.topBorderStyle(), tableStyle.topBorderColor(),
          x, y + height, x + width, y + height);
      renderTableBorderLine(cs, tableStyle.bottomBorderStyle(), tableStyle.bottomBorderColor(),
          x, y, x + width, y);
      renderTableBorderLine(cs, tableStyle.leftBorderStyle(), tableStyle.leftBorderColor(),
          x, y, x, y + height);
      renderTableBorderLine(cs, tableStyle.rightBorderStyle(), tableStyle.rightBorderColor(),
          x + width, y, x + width, y + height);
    }
  }

  /** Backward-compatible single-call version (no overflow). */
  void renderCell(PDPageContentStream cs, @Nullable Cell cell, float x, float y,
      float width, float height, float scaleFactor, @Nullable TableCellStyle tableStyle)
      throws IOException, PdfGenerateException {
    renderBackground(cs, cell, x, y, width, height, tableStyle);
    renderForeground(cs, cell, x, y, width, height, scaleFactor, tableStyle, width);
  }

  // -------------------------------------------------------------------------
  // Text rendering
  // -------------------------------------------------------------------------

  private void renderText(PDPageContentStream cs, Cell cell, String value, float x, float y,
      float width, float height, float scaleFactor, @Nullable Color tableFontColor,
      boolean tableFontBold) throws IOException, PdfGenerateException {

    CellStyle style = cell.getCellStyle();
    Font poiFont = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());

    boolean bold = tableFontBold || poiFont.getBold();
    final boolean italic = poiFont.getItalic();
    final boolean strikeout = poiFont.getStrikeout();
    final boolean underline = poiFont.getUnderline() == Font.U_SINGLE
        || poiFont.getUnderline() == Font.U_SINGLE_ACCOUNTING;
    final boolean doubleUnderline = poiFont.getUnderline() == Font.U_DOUBLE
        || poiFont.getUnderline() == Font.U_DOUBLE_ACCOUNTING;
    final boolean accountingUnderline = poiFont.getUnderline() == Font.U_SINGLE_ACCOUNTING
        || poiFont.getUnderline() == Font.U_DOUBLE_ACCOUNTING;
    final short typeOffset = poiFont.getTypeOffset();
    float fontSize = poiFont.getFontHeightInPoints() * scaleFactor;
    PDType0Font font = fontManager.getFont(bold);

    Color textColor = Color.BLACK;
    if (poiFont instanceof XSSFFont xssfFont) {
      XSSFColor color = xssfFont.getXSSFColor();
      if (color != null) {
        Color c = toAwtColor(color);
        if (c != null) {
          textColor = c;
        }
      }
    }
    if (tableFontColor != null) {
      textColor = tableFontColor;
    }

    if (style.getShrinkToFit() && !style.getWrapText()) {
      float available = width - 2 * CELL_PADDING;
      try {
        float naturalWidth = font.getStringWidth(value) / 1000f * fontSize;
        if (naturalWidth > available && available > 0) {
          fontSize = fontSize * available / naturalWidth;
        }
      } catch (Exception ignored) {
        // keep original size
      }
    }

    boolean superscript = (typeOffset == Font.SS_SUPER);
    boolean subscript = (typeOffset == Font.SS_SUB);
    float effectiveFontSize = (superscript || subscript) ? fontSize * 0.7f : fontSize;

    // Use sTypo metrics (from TTF OS/2 table) for text positioning to match Excel's rendering.
    // PDFBox's font descriptor uses usWinAscent which is inflated for CJK fonts (> 1em for Meiryo).
    float ascent = fontManager.getTypoAscent() / 1000f * effectiveFontSize;
    float descent = fontManager.getTypoDescent() / 1000f * effectiveFontSize;
    float lineHeight = ascent - descent;

    List<String> lines;
    if (style.getWrapText()) {
      float maxLineWidth = width - 2 * CELL_PADDING;
      lines = wrapTextToLines(value, font, effectiveFontSize, maxLineWidth);
    } else {
      lines = List.of(value);
    }

    float totalTextHeight = lines.size() * lineHeight;

    float startY;
    VerticalAlignment vertAlign = getVerticalAlignment(cell, style);
    if (vertAlign == VerticalAlignment.TOP) {
      startY = y + height - CELL_PADDING - ascent;
    } else if (vertAlign == VerticalAlignment.CENTER) {
      startY = y + (height - totalTextHeight) / 2f - descent;
    } else {
      startY = y + CELL_PADDING - descent + totalTextHeight - lineHeight;
    }

    for (String line : lines) {
      if (startY > y + height) {
        startY -= lineHeight;
        continue;
      }
      // Stop when baseline is below cell bottom. We do NOT stop at descender bottom — doing so
      // caused a float edge case where a line whose font height barely exceeds the cell height
      // (by < 1pt) was incorrectly skipped even though its baseline was above the cell bottom.
      if (startY < y) {
        break;
      }
      if (line.isEmpty()) {
        startY -= lineHeight;
        continue;
      }

      float lineY = startY;
      if (superscript) {
        lineY += fontSize * 0.35f;
      } else if (subscript) {
        lineY -= fontSize * 0.15f;
      }

      // Compute text width using per-character font selection (handles fallback fonts).
      float textWidth = fontManager.getStringWidthWithFallback(line, bold, effectiveFontSize);
      final float textX =
          calculateTextX(getHorizontalAlignment(cell, style), cell, x, width, textWidth,
              style.getIndention());

      cs.beginText();
      cs.setNonStrokingColor(textColor);
      if (italic) {
        // Italic uses a text matrix shear; font switching within the matrix is unsupported,
        // so render the whole line with the primary font (best effort for italic+fallback).
        cs.setFont(font, effectiveFontSize);
        cs.setTextMatrix(new Matrix(1, 0, 0.21f, 1, textX, lineY));
        try {
          cs.showText(line);
        } catch (Exception ignored) { // NOPMD
          // Skip lines whose characters are not all in the primary font in italic mode.
        }
      } else {
        cs.newLineAtOffset(textX, lineY);
        // Render each font-segment so that characters not in the primary font use the fallback.
        for (FontManager.TextRun run : fontManager.segmentText(line, bold)) {
          cs.setFont(run.font(), effectiveFontSize);
          cs.showText(run.text());
        }
      }
      cs.endText();

      if (strikeout) {
        float strikeY = lineY + ascent * 0.35f;
        cs.setStrokingColor(textColor);
        cs.setLineWidth(effectiveFontSize / 14f);
        cs.moveTo(textX, strikeY);
        cs.lineTo(textX + textWidth, strikeY);
        cs.stroke();
      }

      if (underline || doubleUnderline) {
        float ulWidth = accountingUnderline ? width - 2 * CELL_PADDING : textWidth;
        float ulX = accountingUnderline ? x + CELL_PADDING : textX;
        float ulY = lineY + descent - 0.5f;
        cs.setStrokingColor(textColor);
        cs.setLineWidth(effectiveFontSize / 14f);
        cs.moveTo(ulX, ulY);
        cs.lineTo(ulX + ulWidth, ulY);
        cs.stroke();
        if (doubleUnderline) {
          cs.moveTo(ulX, ulY - 1.5f);
          cs.lineTo(ulX + ulWidth, ulY - 1.5f);
          cs.stroke();
        }
      }

      startY -= lineHeight;
    }
  }

  private void renderVerticalText(PDPageContentStream cs, Cell cell, String value, float x, float y,
      float width, float height, float scaleFactor) throws IOException, PdfGenerateException {

    CellStyle style = cell.getCellStyle();
    Font poiFont = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());

    boolean bold = poiFont.getBold();
    float fontSize = poiFont.getFontHeightInPoints() * scaleFactor;

    Color textColor = Color.BLACK;
    if (poiFont instanceof XSSFFont xssfFont) {
      XSSFColor color = xssfFont.getXSSFColor();
      if (color != null) {
        Color c = toAwtColor(color);
        if (c != null) {
          textColor = c;
        }
      }
    }

    float ascent = fontManager.getTypoAscent() / 1000f * fontSize;
    float descent = fontManager.getTypoDescent() / 1000f * fontSize;
    float lineHeight = ascent - descent;
    float centerX = x + width / 2f;
    float currentY = y + height - CELL_PADDING - ascent;

    for (int i = 0; i < value.length(); ) {
      if (currentY + descent < y) {
        break;
      }
      int cp = value.codePointAt(i);
      String ch = new String(Character.toChars(cp));
      PDType0Font charFont = fontManager.selectFont(cp, bold);
      float charWidth;
      try {
        charWidth = charFont.getStringWidth(ch) / 1000f * fontSize;
      } catch (Exception ignored) { // NOPMD
        charWidth = fontSize;
      }
      cs.beginText();
      cs.setFont(charFont, fontSize);
      cs.setNonStrokingColor(textColor);
      cs.newLineAtOffset(centerX - charWidth / 2f, currentY);
      cs.showText(ch);
      cs.endText();
      currentY -= lineHeight;
      i += Character.charCount(cp);
    }
  }

  /**
   * Breaks a string into lines that fit within {@code maxWidth} points.
   *
   * <p>Handles both explicit {@code \n} line breaks and automatic wrapping. For CJK
   * text (no word spaces), the break is inserted before the character that would
   * cause the line to exceed {@code maxWidth}. For Latin text with spaces, the break
   * tries to fall on the last space that fits.</p>
   */
  private List<String> wrapTextToLines(String text, PDType0Font font, float fontSize,
      float maxWidth) {
    List<String> result = new ArrayList<>();
    for (String para : text.split("\n", -1)) {
      if (para.isEmpty()) {
        result.add("");
        continue;
      }
      StringBuilder current = new StringBuilder();
      int lastSpaceInCurrent = -1;
      for (int i = 0; i < para.length(); i++) {
        char ch = para.charAt(i);
        current.append(ch);
        if (ch == ' ') {
          lastSpaceInCurrent = current.length() - 1;
        }
        float lineWidth;
        try {
          lineWidth = font.getStringWidth(current.toString()) / 1000f * fontSize;
        } catch (Exception e) {
          lineWidth = 0f;
        }
        if (lineWidth > maxWidth && current.length() > 1) {
          int breakAt = (lastSpaceInCurrent > 0) ? lastSpaceInCurrent : current.length() - 1;
          result.add(current.substring(0, breakAt));
          String rest = current.substring(breakAt).stripLeading();
          current = new StringBuilder(rest);
          lastSpaceInCurrent = -1;
        }
      }
      if (!current.isEmpty()) {
        result.add(current.toString());
      }
    }
    return result;
  }

  private float calculateTextX(HorizontalAlignment align, Cell cell, float x, float width,
      float textWidth, short indent) {
    HorizontalAlignment effective = align;
    if (align == HorizontalAlignment.GENERAL && cell != null) {
      CellType type = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType()
          : cell.getCellType();
      if (type == CellType.NUMERIC || type == CellType.BOOLEAN) {
        effective = HorizontalAlignment.RIGHT;
      }
    }
    float indentPt = indent * INDENT_WIDTH_PT;
    return switch (effective) {
      case CENTER -> x + (width - textWidth) / 2f;
      case RIGHT -> x + width - textWidth - CELL_PADDING - indentPt;
      default -> x + CELL_PADDING + indentPt;
    };
  }

  /**
   * Returns the effective vertical alignment for a cell.
   *
   * <p>Apache POI returns {@link VerticalAlignment#BOTTOM} when the {@code xf}'s
   * {@code applyAlignment} attribute is absent, even if the {@code <alignment>} element
   * has an explicit {@code vertical} attribute. This method works around that by reading
   * the raw {@link org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf}.</p>
   */
  private VerticalAlignment getVerticalAlignment(Cell cell, CellStyle style) {
    VerticalAlignment vertAlign = style.getVerticalAlignment();
    if (vertAlign != VerticalAlignment.BOTTOM) {
      return vertAlign;
    }
    if (style instanceof XSSFCellStyle xssfStyle
        && cell.getSheet().getWorkbook() instanceof XSSFWorkbook xssfWb) {
      int idx = (int) xssfStyle.getIndex();
      var ctXf = xssfWb.getStylesSource().getCellXfAt(idx);
      var rawAlign = ctXf.getAlignment();
      if (rawAlign != null && rawAlign.isSetVertical()) {
        var sv = rawAlign.getVertical();
        if (sv == STVerticalAlignment.TOP) {
          return VerticalAlignment.TOP;
        }
        if (sv == STVerticalAlignment.CENTER) {
          return VerticalAlignment.CENTER;
        }
        if (sv == STVerticalAlignment.JUSTIFY) {
          return VerticalAlignment.JUSTIFY;
        }
        if (sv == STVerticalAlignment.DISTRIBUTED) {
          return VerticalAlignment.DISTRIBUTED;
        }
      }
    }
    return vertAlign;
  }

  /**
   * Returns the effective horizontal alignment for a cell.
   *
   * <p>Same workaround as {@link #getVerticalAlignment}: when {@code applyAlignment} is absent
   * Apache POI returns {@link HorizontalAlignment#GENERAL}, ignoring an explicit
   * {@code horizontal} attribute.</p>
   */
  HorizontalAlignment getHorizontalAlignment(Cell cell, CellStyle style) {
    HorizontalAlignment halign = style.getAlignment();
    if (halign != HorizontalAlignment.GENERAL) {
      return halign;
    }
    if (style instanceof XSSFCellStyle xssfStyle
        && cell.getSheet().getWorkbook() instanceof XSSFWorkbook xssfWb) {
      int idx = (int) xssfStyle.getIndex();
      var ctXf = xssfWb.getStylesSource().getCellXfAt(idx);
      var rawAlign = ctXf.getAlignment();
      if (rawAlign != null && rawAlign.isSetHorizontal()) {
        var sh = rawAlign.getHorizontal();
        if (sh == STHorizontalAlignment.LEFT) {
          return HorizontalAlignment.LEFT;
        }
        if (sh == STHorizontalAlignment.RIGHT) {
          return HorizontalAlignment.RIGHT;
        }
        if (sh == STHorizontalAlignment.CENTER) {
          return HorizontalAlignment.CENTER;
        }
        if (sh == STHorizontalAlignment.FILL) {
          return HorizontalAlignment.FILL;
        }
        if (sh == STHorizontalAlignment.JUSTIFY) {
          return HorizontalAlignment.JUSTIFY;
        }
        if (sh == STHorizontalAlignment.CENTER_CONTINUOUS) {
          return HorizontalAlignment.CENTER_SELECTION;
        }
        if (sh == STHorizontalAlignment.DISTRIBUTED) {
          return HorizontalAlignment.DISTRIBUTED;
        }
      }
    }
    return halign;
  }

  // -------------------------------------------------------------------------
  // Border rendering
  // -------------------------------------------------------------------------

  void renderBorderLine(PDPageContentStream cs, BorderStyle borderStyle, Color color,
      float x1, float y1, float x2, float y2) throws IOException {
    if (borderStyle == BorderStyle.NONE) {
      return;
    }
    float[] dash = getDashPattern(borderStyle);
    cs.setStrokingColor(color);
    cs.setLineWidth(getBorderLineWidth(borderStyle));
    if (dash != null) {
      cs.setLineDashPattern(dash, 0);
    }
    cs.moveTo(x1, y1);
    cs.lineTo(x2, y2);
    cs.stroke();
    if (dash != null) {
      cs.setLineDashPattern(new float[] {}, 0);
    }
  }

  private void renderBorders(PDPageContentStream cs, CellStyle style, Sheet sheet, float x, float y,
      float width, float height) throws IOException {
    XSSFCellStyle xssfStyle = (style instanceof XSSFCellStyle s) ? s : null;

    renderBorderLine(cs, style.getBorderTop(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getTopBorderXSSFColor() : null),
        x, y + height, x + width, y + height);
    renderBorderLine(cs, style.getBorderBottom(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getBottomBorderXSSFColor() : null),
        x, y, x + width, y);
    renderBorderLine(cs, style.getBorderLeft(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getLeftBorderXSSFColor() : null),
        x, y, x, y + height);
    renderBorderLine(cs, style.getBorderRight(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getRightBorderXSSFColor() : null),
        x + width, y, x + width, y + height);
    if (xssfStyle != null) {
      renderDiagonalBorders(cs, xssfStyle, sheet, x, y, width, height);
    }
  }

  private void renderDiagonalBorders(PDPageContentStream cs, XSSFCellStyle style, Sheet sheet,
      float x, float y, float width, float height) throws IOException {
    if (!(sheet.getWorkbook() instanceof XSSFWorkbook xssfWb)) {
      return;
    }
    int borderId = (int) style.getCoreXf().getBorderId();
    XSSFCellBorder cellBorder = xssfWb.getStylesSource().getBorderAt(borderId);
    BorderStyle diagStyle = cellBorder.getBorderStyle(XSSFCellBorder.BorderSide.DIAGONAL);
    if (diagStyle == BorderStyle.NONE) {
      return;
    }
    Color color = xssfColorToAwt(cellBorder.getBorderColor(XSSFCellBorder.BorderSide.DIAGONAL));
    CTBorder ctBorder = cellBorder.getCTBorder();
    if (ctBorder.getDiagonalDown()) {
      renderBorderLine(cs, diagStyle, color, x, y + height, x + width, y);
    }
    if (ctBorder.getDiagonalUp()) {
      renderBorderLine(cs, diagStyle, color, x, y, x + width, y + height);
    }
  }

  private void renderTableBorderLine(PDPageContentStream cs, @Nullable BorderStyle borderStyle,
      @Nullable Color color, float x1, float y1, float x2, float y2) throws IOException {
    if (borderStyle == null || borderStyle == BorderStyle.NONE) {
      return;
    }
    renderBorderLine(cs, borderStyle, color != null ? color : Color.BLACK, x1, y1, x2, y2);
  }

  private float getBorderLineWidth(BorderStyle style) {
    return switch (style) {
      case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT -> 1.0f;
      case THICK -> 1.5f;
      default -> 0.5f;
    };
  }

  private float @Nullable [] getDashPattern(BorderStyle style) {
    return switch (style) {
      case DASHED, MEDIUM_DASHED -> new float[] {4f, 3f};
      case DOTTED -> new float[] {1f, 2f};
      case DASH_DOT, MEDIUM_DASH_DOT, SLANTED_DASH_DOT -> new float[] {4f, 2f, 1f, 2f};
      case DASH_DOT_DOT, MEDIUM_DASH_DOT_DOT -> new float[] {4f, 2f, 1f, 2f, 1f, 2f};
      default -> null;
    };
  }

  // -------------------------------------------------------------------------
  // Color utilities
  // -------------------------------------------------------------------------

  Color xssfColorToAwt(@Nullable XSSFColor xssfColor) {
    if (xssfColor == null) {
      return Color.BLACK;
    }
    Color c = toAwtColor(xssfColor);
    return (c != null) ? c : Color.BLACK;
  }

  @Nullable Color toAwtColor(org.apache.poi.ss.usermodel.Color poiColor) {
    if (poiColor instanceof XSSFColor xssfColor) {
      byte[] rgb = xssfColor.getRGBWithTint();
      if (rgb == null) {
        rgb = xssfColor.getRGB();
      }
      if (rgb != null && rgb.length == 3) {
        return new Color(Byte.toUnsignedInt(rgb[0]), Byte.toUnsignedInt(rgb[1]),
            Byte.toUnsignedInt(rgb[2]));
      }
    }
    return null;
  }

  /**
   * Converts a POI {@link org.apache.poi.ss.usermodel.Color} to {@link Color}, resolving
   * theme-based colours via the workbook's {@link org.apache.poi.xssf.model.ThemesTable} when
   * the normal {@link XSSFColor#getRGBWithTint()} cannot resolve the colour on its own.
   */
  @Nullable Color poiColorToAwt(org.apache.poi.ss.usermodel.Color color) {
    if (!(color instanceof XSSFColor xssfColor)) {
      return null;
    }
    byte[] rgb = xssfColor.getRGBWithTint();
    if (rgb == null && currentWorkbook != null && xssfColor.getCTColor().isSetTheme()) {
      var theme = currentWorkbook.getStylesSource().getTheme();
      if (theme != null) {
        theme.inheritFromThemeAsRequired(xssfColor);
        rgb = xssfColor.getRGBWithTint();
      }
    }
    if (rgb == null) {
      return null;
    }
    return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
  }

  @Nullable Color getBackgroundColor(CellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return toAwtColor(style.getFillForegroundColorColor());
  }

  // -------------------------------------------------------------------------
  // Table style resolution
  // -------------------------------------------------------------------------

  List<TableRenderInfo> collectTableRenderInfos(Sheet sheet, Workbook workbook) {
    if (!(sheet instanceof XSSFSheet xssfSheet)) {
      return List.of();
    }
    if (!(workbook instanceof XSSFWorkbook xssfWb)) {
      return List.of();
    }

    List<XSSFTable> tables = xssfSheet.getTables();
    if (tables.isEmpty()) {
      return List.of();
    }

    var stylesSource = xssfWb.getStylesSource();
    List<TableRenderInfo> result = new ArrayList<>();

    for (XSSFTable table : tables) {
      var rawInfo = table.getStyle();
      if (rawInfo == null) {
        continue;
      }

      String name = rawInfo.getName();
      if (name == null || name.isBlank()) {
        continue;
      }

      TableStyle tableStyle = stylesSource.getTableStyle(name);
      if (tableStyle == null) {
        continue;
      }

      var areaRef = table.getArea();
      if (areaRef == null) {
        continue;
      }
      var firstCell = areaRef.getFirstCell();
      var lastCell = areaRef.getLastCell();
      CellRangeAddress area = new CellRangeAddress(
          firstCell.getRow(), lastCell.getRow(),
          firstCell.getCol(), lastCell.getCol());

      DifferentialStyleProvider wholeTable = tableStyle.getStyle(TableStyleType.wholeTable);
      DifferentialStyleProvider headerRow = tableStyle.getStyle(TableStyleType.headerRow);
      DifferentialStyleProvider firstRowStripe = tableStyle.getStyle(TableStyleType.firstRowStripe);
      DifferentialStyleProvider secondRowStripe =
          tableStyle.getStyle(TableStyleType.secondRowStripe);
      DifferentialStyleProvider firstColumn = tableStyle.getStyle(TableStyleType.firstColumn);
      DifferentialStyleProvider lastColumn = tableStyle.getStyle(TableStyleType.lastColumn);

      int firstStripeSize =
          (firstRowStripe != null && firstRowStripe.getStripeSize() > 0)
              ? firstRowStripe.getStripeSize() : 1;
      int secondStripeSize =
          (secondRowStripe != null && secondRowStripe.getStripeSize() > 0)
              ? secondRowStripe.getStripeSize() : 1;

      result.add(new TableRenderInfo(
          area, area.getFirstRow(), area.getFirstRow(),
          rawInfo.isShowRowStripes(), rawInfo.isShowFirstColumn(), rawInfo.isShowLastColumn(),
          wholeTable, headerRow, firstRowStripe, secondRowStripe, firstColumn, lastColumn,
          firstStripeSize, secondStripeSize));
    }
    return result;
  }

  @Nullable TableCellStyle getTableCellStyle(List<TableRenderInfo> tables, int row, int col) {
    for (TableRenderInfo info : tables) {
      if (!info.area().isInRange(row, col)) {
        continue;
      }

      final boolean isHeader = row >= info.headerFirstRow() && row <= info.headerLastRow();
      final boolean isLastRow = row == info.area().getLastRow();
      final boolean isFirstCol = col == info.area().getFirstColumn();
      final boolean isLastCol = col == info.area().getLastColumn();
      final boolean isFirstDataRow = row == info.headerLastRow() + 1;

      List<DifferentialStyleProvider> providers = new ArrayList<>();
      if (info.wholeTable() != null) {
        providers.add(info.wholeTable());
      }

      if (!isHeader && info.showRowStripes() && info.firstRowStripe() != null) {
        int dataRow = row - info.headerLastRow() - 1;
        int cycle = info.firstStripeSize() + info.secondStripeSize();
        boolean inFirstStripe = (dataRow % cycle) < info.firstStripeSize();
        if (inFirstStripe) {
          providers.add(info.firstRowStripe());
        } else if (info.secondRowStripe() != null) {
          providers.add(info.secondRowStripe());
        }
      }

      if (info.showFirstColumn() && isFirstCol && info.firstColumn() != null) {
        providers.add(info.firstColumn());
      }
      if (info.showLastColumn() && isLastCol && info.lastColumn() != null) {
        providers.add(info.lastColumn());
      }
      if (isHeader && info.headerRow() != null) {
        providers.add(info.headerRow());
      }

      Color fill = null;
      Color fontColor = null;
      boolean fontBold = false;
      BorderStyle topStyle = null;
      Color topColor = null;
      BorderStyle bottomStyle = null;
      Color bottomColor = null;
      BorderStyle leftStyle = null;
      Color leftColor = null;
      BorderStyle rightStyle = null;
      Color rightColor = null;

      for (var p : providers) {
        var pf = p.getPatternFormatting();
        if (pf != null) {
          Color c = poiColorToAwt(pf.getFillForegroundColorColor());
          if (c != null) {
            fill = c;
          }
        }

        var ff = p.getFontFormatting();
        if (ff != null) {
          Color fc = poiColorToAwt(ff.getFontColor());
          if (fc != null) {
            fontColor = fc;
          }
          if (ff.isBold()) {
            fontBold = true;
          }
        }

        var bf = p.getBorderFormatting();
        if (bf != null) {
          BorderStyle hs = bf.getBorderHorizontal();
          Color hc = poiColorToAwt(bf.getHorizontalBorderColorColor());
          if (hs != null && hs != BorderStyle.NONE) {
            bottomStyle = hs;
            bottomColor = hc;
          }
          if (isHeader || isFirstDataRow) {
            BorderStyle ts = bf.getBorderTop();
            Color tc = poiColorToAwt(bf.getTopBorderColorColor());
            if (ts != null && ts != BorderStyle.NONE) {
              topStyle = ts;
              topColor = tc;
            }
          }
          if (isLastRow) {
            BorderStyle bs = bf.getBorderBottom();
            Color bc = poiColorToAwt(bf.getBottomBorderColorColor());
            if (bs != null && bs != BorderStyle.NONE) {
              bottomStyle = bs;
              bottomColor = bc;
            }
          }
          if (isFirstCol) {
            BorderStyle ls = bf.getBorderLeft();
            Color lc = poiColorToAwt(bf.getLeftBorderColorColor());
            if (ls != null && ls != BorderStyle.NONE) {
              leftStyle = ls;
              leftColor = lc;
            }
          }
          if (isLastCol) {
            BorderStyle rs = bf.getBorderRight();
            Color rc = poiColorToAwt(bf.getRightBorderColorColor());
            if (rs != null && rs != BorderStyle.NONE) {
              rightStyle = rs;
              rightColor = rc;
            }
          }
        }
      }

      boolean isStructural = isHeader
          || (info.showFirstColumn() && isFirstCol)
          || (info.showLastColumn() && isLastCol);

      if (isHeader && fontColor == null && fill != null) {
        double lum = (0.2126 * fill.getRed() + 0.7152 * fill.getGreen()
            + 0.0722 * fill.getBlue()) / 255.0;
        if (lum < 0.5) {
          fontColor = Color.WHITE;
        }
      }
      Color effectiveFontColor = isStructural ? fontColor : null;
      boolean effectiveFontBold = isStructural && fontBold;

      return new TableCellStyle(fill,
          topStyle, topColor, bottomStyle, bottomColor,
          leftStyle, leftColor, rightStyle, rightColor,
          effectiveFontColor, effectiveFontBold);
    }
    return null;
  }
}
