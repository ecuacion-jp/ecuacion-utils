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

import java.awt.Color;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import jp.ecuacion.util.pdfbox.excel.exception.PdfGenerateException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.TextAlign;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/**
 * Renders an Excel sheet into one or more PDF pages in a {@link PDDocument}.
 *
 * <p>Rendering respects the sheet's print area, page setup (paper size, orientation, margins,
 * scale), manual page breaks, cell background colors, borders, and text styles.</p>
 */
public class SheetRenderer {

  /** Pixels-to-points conversion factor (96 DPI screen to 72 DPI points). */
  private static final float PX_TO_PT = 72f / 96f;

  /** Horizontal padding inside a cell, in points. */
  private static final float CELL_PADDING = 2f;

  private final PDDocument document;
  private final FontManager fontManager;
  private final DataFormatter dataFormatter = new DataFormatter();

  /**
   * Constructs a {@code SheetRenderer}.
   *
   * @param document the target PDF document
   * @param fontManager the font manager providing embedded fonts
   */
  public SheetRenderer(PDDocument document, FontManager fontManager) {
    this.document = document;
    this.fontManager = fontManager;
  }

  /**
   * Renders the specified sheet into the PDF document.
   *
   * @param workbook the workbook containing the sheet
   * @param sheetIndex 0-based index of the sheet
   * @throws IOException if a PDF I/O error occurs
   * @throws PdfGenerateException if the sheet cannot be rendered
   */
  public void render(Workbook workbook, int sheetIndex) throws IOException, PdfGenerateException {
    Sheet sheet = workbook.getSheetAt(sheetIndex);

    int[] bounds = getPrintAreaBounds(workbook, sheet, sheetIndex);
    int firstRow = bounds[0];
    int lastRow = bounds[1];
    int firstCol = bounds[2];
    int lastCol = bounds[3];

    PrintSetup ps = sheet.getPrintSetup();
    PDRectangle basePageSize = getPageSize(ps);
    PDRectangle pageSize = ps.getLandscape()
        ? new PDRectangle(basePageSize.getHeight(), basePageSize.getWidth())
        : basePageSize;

    float leftMargin = (float) (sheet.getMargin(PageMargin.LEFT) * 72);
    float rightMargin = (float) (sheet.getMargin(PageMargin.RIGHT) * 72);
    float topMargin = (float) (sheet.getMargin(PageMargin.TOP) * 72);
    float bottomMargin = (float) (sheet.getMargin(PageMargin.BOTTOM) * 72);
    float printableWidth = pageSize.getWidth() - leftMargin - rightMargin;
    float printableHeight = pageSize.getHeight() - topMargin - bottomMargin;

    float scaleFactor = computeScaleFactor(sheet, ps, firstRow, lastRow, firstCol, lastCol,
        printableWidth, printableHeight);

    float[] colWidths = new float[lastCol - firstCol + 1];
    float naturalColTotal = 0f;
    for (int c = firstCol; c <= lastCol; c++) {
      float natural = getColumnNaturalWidthInPt(sheet, c);
      naturalColTotal += natural;
      colWidths[c - firstCol] = natural * scaleFactor;
    }

    float[] rowHeights = new float[lastRow - firstRow + 1];
    for (int r = firstRow; r <= lastRow; r++) {
      Row row = sheet.getRow(r);
      float h = (row != null) ? row.getHeightInPoints() : sheet.getDefaultRowHeightInPoints();
      rowHeights[r - firstRow] = h * scaleFactor;
    }

    Map<String, CellRangeAddress> mergedRegionMap = buildMergedRegionMap(sheet);

    List<int[]> rowPages = buildPageRanges(firstRow, lastRow, sheet.getRowBreaks(),
        rowHeights, printableHeight);
    // When no manual column breaks are defined and the natural (unscaled) column total
    // fits within the printable width, treat the sheet as single-column-page. An explicit
    // scale > 1 may push the scaled total slightly over the boundary, but the intent is
    // to print all columns on one page — consistent with Excel's behavior.
    float colPageWidth = (sheet.getColumnBreaks().length == 0
        && naturalColTotal <= printableWidth)
        ? Float.MAX_VALUE : printableWidth;
    List<int[]> colPages = buildPageRanges(firstCol, lastCol, sheet.getColumnBreaks(),
        colWidths, colPageWidth);

    for (int[] colPage : colPages) {
      for (int[] rowPage : rowPages) {
        renderPage(sheet, pageSize, leftMargin, topMargin,
            rowPage[0], rowPage[1], colPage[0], colPage[1],
            firstRow, firstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap);
      }
    }
  }

  // -------------------------------------------------------------------------
  // Page setup helpers
  // -------------------------------------------------------------------------

  private int[] getPrintAreaBounds(Workbook workbook, Sheet sheet, int sheetIndex)
      throws PdfGenerateException {
    String printArea = workbook.getPrintArea(sheetIndex);

    if (printArea != null && !printArea.isBlank()) {
      // Format: "SheetName!$A$1:$H$30" – strip sheet name and $ signs
      String ref = printArea.contains("!")
          ? printArea.substring(printArea.indexOf('!') + 1) : printArea;
      ref = ref.replace("$", "");
      String[] parts = ref.split(":");
      if (parts.length == 2) {
        int firstRow = cellRefToRow(parts[0]);
        int firstCol = cellRefToCol(parts[0]);
        int lastRow = cellRefToRow(parts[1]);
        int lastCol = cellRefToCol(parts[1]);
        return new int[] {firstRow, lastRow, firstCol, lastCol};
      }
    }

    // Fall back to the used range of the sheet
    int firstRow = sheet.getFirstRowNum();
    int lastRow = sheet.getLastRowNum();
    int firstCol = Integer.MAX_VALUE;
    int lastCol = 0;
    for (Row row : sheet) {
      if (row.getFirstCellNum() >= 0) {
        firstCol = Math.min(firstCol, row.getFirstCellNum());
      }
      if (row.getLastCellNum() > 0) {
        lastCol = Math.max(lastCol, row.getLastCellNum() - 1);
      }
    }
    if (firstCol == Integer.MAX_VALUE) {
      throw new PdfGenerateException(
          "Sheet '" + sheet.getSheetName() + "' has no print area and no data.");
    }
    return new int[] {firstRow, lastRow, firstCol, lastCol};
  }

  /** Returns the 0-based row index from a cell reference like "A1" or "B10". */
  private int cellRefToRow(String ref) {
    // Extract digits at the end
    int i = ref.length() - 1;
    while (i >= 0 && Character.isDigit(ref.charAt(i))) {
      i--;
    }
    return Integer.parseInt(ref.substring(i + 1)) - 1;
  }

  /** Returns the 0-based column index from a cell reference like "A1" or "AB3". */
  private int cellRefToCol(String ref) {
    int col = 0;
    for (char ch : ref.toCharArray()) {
      if (Character.isLetter(ch)) {
        col = col * 26 + (Character.toUpperCase(ch) - 'A' + 1);
      }
    }
    return col - 1;
  }

  private PDRectangle getPageSize(PrintSetup ps) {
    // Common Excel paper size codes
    return switch (ps.getPaperSize()) {
      case PrintSetup.LETTER_PAPERSIZE -> PDRectangle.LETTER;
      case PrintSetup.A5_PAPERSIZE -> PDRectangle.A5;
      default -> PDRectangle.A4; // includes PrintSetup.A4_PAPERSIZE
    };
  }

  /**
   * Computes the scale factor for rendering a sheet.
   *
   * <p>When the page setup XML contains an explicit {@code scale} attribute the value is used
   * directly. Otherwise the sheet is treated as "fit to page": the content is scaled down so
   * that all columns and rows fit within the printable area in one pass.</p>
   *
   * @param sheet the sheet being rendered
   * @param ps the print setup of the sheet
   * @param firstRow first row index of the print area
   * @param lastRow last row index of the print area
   * @param firstCol first column index of the print area
   * @param lastCol last column index of the print area
   * @param printableWidth available width in points (page width minus left and right margins)
   * @param printableHeight available height in points (page height minus top and bottom margins)
   * @return the scale factor to apply to column widths and row heights
   */
  private float computeScaleFactor(Sheet sheet, PrintSetup ps,
      int firstRow, int lastRow, int firstCol, int lastCol,
      float printableWidth, float printableHeight) {
    // For XSSF sheets, check whether the scale attribute is explicitly present in the XML.
    if (sheet instanceof XSSFSheet xssfSheet
        && xssfSheet.getCTWorksheet().isSetPageSetup()) {
      var ctPageSetup = xssfSheet.getCTWorksheet().getPageSetup();
      if (ctPageSetup.isSetScale()) {
        long s = ctPageSetup.getScale();
        return (s > 0 && s <= 400) ? s / 100f : 1f;
      }
    }
    // No explicit scale attribute: treat the sheet as "fit to page".
    // Compute the natural (unscaled) total column width and row height, then derive
    // the minimum scale that makes the content fit within the printable area.
    float naturalColTotal = 0f;
    for (int c = firstCol; c <= lastCol; c++) {
      naturalColTotal += getColumnNaturalWidthInPt(sheet, c);
    }
    float naturalRowTotal = 0f;
    for (int r = firstRow; r <= lastRow; r++) {
      Row row = sheet.getRow(r);
      float h = (row != null) ? row.getHeightInPoints() : sheet.getDefaultRowHeightInPoints();
      naturalRowTotal += h;
    }
    float fitScale = 1.0f;
    if (naturalColTotal > printableWidth) {
      fitScale = Math.min(fitScale, printableWidth / naturalColTotal);
    }
    if (naturalRowTotal > printableHeight) {
      fitScale = Math.min(fitScale, printableHeight / naturalRowTotal);
    }
    return fitScale;
  }

  /**
   * Returns the natural (unscaled) width of the specified column in points.
   *
   * <p>For XSSF sheets whose column has no explicit custom width, Apache POI falls back to its
   * built-in default of 8 characters rather than reading the {@code defaultColWidth} attribute
   * from the sheet XML. This method corrects that by reading the XML attribute directly and
   * computing the column width as {@code defaultColWidth × 7px/char × PX_TO_PT}, which matches
   * Excel's rendering for sheets that use a narrow default column width.</p>
   *
   * @param sheet the sheet
   * @param col the 0-based column index
   * @return column width in points
   */
  private float getColumnNaturalWidthInPt(Sheet sheet, int col) {
    if (sheet instanceof XSSFSheet xssfSheet) {
      var ws = xssfSheet.getCTWorksheet();
      // Check whether this column has an explicit custom width in the <cols> element.
      for (var ctCols : ws.getColsArray()) {
        for (CTCol ctCol : ctCols.getColList()) {
          if (ctCol.isSetCustomWidth() && ctCol.getCustomWidth()
              && ctCol.getMin() <= col + 1 && col + 1 <= ctCol.getMax()) {
            // Explicit width: trust POI's calculation.
            return sheet.getColumnWidthInPixels(col) * PX_TO_PT;
          }
        }
      }
      // No explicit width: use the actual defaultColWidth from the XML instead of POI's
      // built-in fallback of 8 characters.
      // Note: isSetDefaultColWidth() is not used because XMLBeans may fail to detect
      // the attribute as explicitly set; getDefaultColWidth() returns the schema default
      // (8) when not set, which produces the same result as POI's fallback.
      if (ws.isSetSheetFormatPr()) {
        double dcw = ws.getSheetFormatPr().getDefaultColWidth();
        if (dcw > 0) {
          // 7px is the standard max-digit-width for the default Calibri 11pt font at 96 dpi.
          return (float) (dcw * 7.0 * PX_TO_PT);
        }
      }
    }
    return sheet.getColumnWidthInPixels(col) * PX_TO_PT;
  }

  // -------------------------------------------------------------------------
  // Page break calculation
  // -------------------------------------------------------------------------

  /**
   * Splits the range [{@code first}, {@code last}] into pages.
   * Pages are split at manual breaks or when the cumulative size exceeds {@code maxSize}.
   */
  private List<int[]> buildPageRanges(int first, int last, int[] manualBreaks,
      float[] sizes, float maxSize) {
    Set<Integer> breakSet = new HashSet<>();
    for (int b : manualBreaks) {
      breakSet.add(b);
    }

    List<int[]> pages = new ArrayList<>();
    int pageStart = first;
    float currentSize = 0f;

    for (int i = first; i <= last; i++) {
      float size = sizes[i - first];

      // Automatic break: adding this row/col would exceed the page
      if (currentSize + size > maxSize && i > pageStart) {
        pages.add(new int[] {pageStart, i - 1});
        pageStart = i;
        currentSize = size;
      } else {
        currentSize += size;
      }

      // Manual break after index i
      if (breakSet.contains(i) && i < last) {
        pages.add(new int[] {pageStart, i});
        pageStart = i + 1;
        currentSize = 0f;
      }
    }

    if (pageStart <= last) {
      pages.add(new int[] {pageStart, last});
    }
    return pages;
  }

  // -------------------------------------------------------------------------
  // Merged region map
  // -------------------------------------------------------------------------

  private Map<String, CellRangeAddress> buildMergedRegionMap(Sheet sheet) {
    Map<String, CellRangeAddress> map = new HashMap<>();
    for (CellRangeAddress region : sheet.getMergedRegions()) {
      for (int r = region.getFirstRow(); r <= region.getLastRow(); r++) {
        for (int c = region.getFirstColumn(); c <= region.getLastColumn(); c++) {
          map.put(r + "," + c, region);
        }
      }
    }
    return map;
  }

  // -------------------------------------------------------------------------
  // Page rendering
  // -------------------------------------------------------------------------

  private void renderPage(Sheet sheet, PDRectangle pageSize,
      float leftMargin, float topMargin,
      int firstPageRow, int lastPageRow, int firstPageCol, int lastPageCol,
      int printFirstRow, int printFirstCol,
      float[] rowHeights, float[] colWidths, float scaleFactor,
      Map<String, CellRangeAddress> mergedRegionMap) throws IOException {

    PDPage page = new PDPage(pageSize);
    document.addPage(page);

    try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
      float currentY = pageSize.getHeight() - topMargin;

      for (int r = firstPageRow; r <= lastPageRow; r++) {
        float rowHeight = rowHeights[r - printFirstRow];
        float currentX = leftMargin;

        Row row = sheet.getRow(r);

        for (int c = firstPageCol; c <= lastPageCol; c++) {
          float colWidth = colWidths[c - printFirstCol];

          CellRangeAddress region = mergedRegionMap.get(r + "," + c);

          // Skip cells that are part of a merged region but not the top-left cell
          if (region != null
              && (region.getFirstRow() != r || region.getFirstColumn() != c)) {
            currentX += colWidth;
            continue;
          }

          // Compute actual width/height, summing across merged span
          float cellWidth = colWidth;
          float cellHeight = rowHeight;

          if (region != null) {
            cellWidth = 0f;
            for (int mc = region.getFirstColumn(); mc <= region.getLastColumn(); mc++) {
              if (mc >= firstPageCol && mc <= lastPageCol) {
                cellWidth += colWidths[mc - printFirstCol];
              }
            }
            cellHeight = 0f;
            for (int mr = region.getFirstRow(); mr <= region.getLastRow(); mr++) {
              if (mr >= firstPageRow && mr <= lastPageRow) {
                cellHeight += rowHeights[mr - printFirstRow];
              }
            }
          }

          Cell cell = (row != null) ? row.getCell(c) : null;
          // PDF y origin is bottom-left; currentY is the top of the current row
          float cellBottomY = currentY - cellHeight;

          renderCell(cs, cell, currentX, cellBottomY, cellWidth, cellHeight, scaleFactor);

          // For merged cells, right/bottom borders come from the boundary cells, not the
          // top-left cell whose style may lack those borders.
          if (region != null && region.getLastColumn() > region.getFirstColumn()) {
            Cell rightBoundary = (row != null) ? row.getCell(region.getLastColumn()) : null;
            if (rightBoundary != null) {
              CellStyle rightStyle = rightBoundary.getCellStyle();
              XSSFCellStyle rxssf = (rightStyle instanceof XSSFCellStyle s) ? s : null;
              renderBorderLine(cs, rightStyle.getBorderRight(),
                  xssfColorToAwt(rxssf != null ? rxssf.getRightBorderXSSFColor() : null),
                  currentX + cellWidth, cellBottomY,
                  currentX + cellWidth, cellBottomY + cellHeight);
            }
          }
          if (region != null && region.getLastRow() > region.getFirstRow()) {
            Row lastMergeRow = sheet.getRow(region.getLastRow());
            Cell bottomBoundary = (lastMergeRow != null)
                ? lastMergeRow.getCell(region.getFirstColumn()) : null;
            if (bottomBoundary != null) {
              CellStyle bottomStyle = bottomBoundary.getCellStyle();
              XSSFCellStyle bxssf = (bottomStyle instanceof XSSFCellStyle s) ? s : null;
              renderBorderLine(cs, bottomStyle.getBorderBottom(),
                  xssfColorToAwt(bxssf != null ? bxssf.getBottomBorderXSSFColor() : null),
                  currentX, cellBottomY, currentX + cellWidth, cellBottomY);
            }
          }

          currentX += colWidth;
        }

        currentY -= rowHeight;
      }

      // Shapes are rendered on top of cells
      renderShapes(sheet, cs, pageSize, leftMargin, topMargin,
          firstPageRow, lastPageRow, firstPageCol, lastPageCol,
          printFirstRow, printFirstCol, rowHeights, colWidths, scaleFactor);
    }
  }

  // -------------------------------------------------------------------------
  // Shape rendering
  // -------------------------------------------------------------------------

  private void renderShapes(Sheet sheet, PDPageContentStream cs,
      PDRectangle pageSize, float leftMargin, float topMargin,
      int firstPageRow, int lastPageRow, int firstPageCol, int lastPageCol,
      int printFirstRow, int printFirstCol,
      float[] rowHeights, float[] colWidths, float scaleFactor) throws IOException {

    if (!(sheet instanceof XSSFSheet xssfSheet)) {
      return;
    }
    XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
    if (drawing == null) {
      return;
    }

    // Y offset (in Excel coords) from print area top to this page's top
    float pageTopExcelY = 0f;
    for (int r = printFirstRow; r < firstPageRow; r++) {
      pageTopExcelY += rowHeights[r - printFirstRow];
    }
    float pageBottomExcelY = pageTopExcelY;
    for (int r = firstPageRow; r <= lastPageRow; r++) {
      pageBottomExcelY += rowHeights[r - printFirstRow];
    }

    // X offset (in Excel coords) from print area left to this page's left
    float pageLeftExcelX = 0f;
    for (int c = printFirstCol; c < firstPageCol; c++) {
      pageLeftExcelX += colWidths[c - printFirstCol];
    }

    for (XSSFShape shape : drawing.getShapes()) {
      if (!(shape instanceof XSSFSimpleShape simpleShape)) {
        continue;
      }
      if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor)) {
        continue;
      }

      // Shape bounds in Excel coordinates (from print area start).
      // For two-cell anchors (row2 or col2 > 0) derive the size from the second anchor so
      // that the shape scales correctly with the row/column layout.
      // For one-cell anchors (row2 == col2 == 0) fall back to xfrm.ext, scaled uniformly.
      float shapeTopY =
          shapeExcelY(anchor.getRow1(), anchor.getDy1(), rowHeights, printFirstRow);
      float shapeLeftX =
          shapeExcelX(anchor.getCol1(), anchor.getDx1(), colWidths, printFirstCol);
      float shapeBottomY;
      float shapeRightX;
      if (anchor.getRow2() > 0 || anchor.getCol2() > 0) {
        shapeBottomY = shapeExcelY(anchor.getRow2(), anchor.getDy2(), rowHeights, printFirstRow);
        shapeRightX = shapeExcelX(anchor.getCol2(), anchor.getDx2(), colWidths, printFirstCol);
      } else {
        long extCx = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCx();
        long extCy = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCy();
        shapeBottomY = shapeTopY + extCy / 12700f * scaleFactor;
        shapeRightX = shapeLeftX + extCx / 12700f * scaleFactor;
      }

      // Render only shapes whose top falls within this page
      if (shapeTopY < pageTopExcelY || shapeTopY >= pageBottomExcelY) {
        continue;
      }

      float relTop = shapeTopY - pageTopExcelY;
      float relBottom = shapeBottomY - pageTopExcelY;
      float relLeft = shapeLeftX - pageLeftExcelX;

      float pdfShapeBottom = pageSize.getHeight() - topMargin - relBottom;
      float pdfShapeHeight = relBottom - relTop;
      float pdfShapeLeft = leftMargin + relLeft;
      float pdfShapeWidth = shapeRightX - shapeLeftX;

      renderShape(cs, simpleShape, pdfShapeLeft, pdfShapeBottom,
          pdfShapeWidth, pdfShapeHeight, scaleFactor);
    }
  }

  /** Converts a shape anchor's row + EMU offset to an Excel-coordinate Y value. */
  private float shapeExcelY(int row, int dyEmu, float[] rowHeights, int printFirstRow) {
    float y = 0f;
    int limit = Math.min(row, printFirstRow + rowHeights.length);
    for (int r = printFirstRow; r < limit; r++) {
      y += rowHeights[r - printFirstRow];
    }
    return y + dyEmu / 12700f;
  }

  /** Converts a shape anchor's column + EMU offset to an Excel-coordinate X value. */
  private float shapeExcelX(int col, int dxEmu, float[] colWidths, int printFirstCol) {
    float x = 0f;
    int limit = Math.min(col, printFirstCol + colWidths.length);
    for (int c = printFirstCol; c < limit; c++) {
      x += colWidths[c - printFirstCol];
    }
    return x + dxEmu / 12700f;
  }

  private void renderShape(PDPageContentStream cs, XSSFSimpleShape shape,
      float x, float y, float width, float height, float scaleFactor) throws IOException {

    Color fillColor = getShapeFillColor(shape);
    Color lineColor = getShapeLineColor(shape);
    float lineWidth = getShapeLineWidth(shape);
    boolean hasLine = lineColor != null && lineWidth > 0f;

    if (fillColor != null && hasLine) {
      cs.setNonStrokingColor(fillColor);
      cs.setStrokingColor(lineColor);
      cs.setLineWidth(lineWidth);
      appendShapePath(cs, shape, x, y, width, height);
      cs.fillAndStroke();
    } else if (fillColor != null) {
      cs.setNonStrokingColor(fillColor);
      appendShapePath(cs, shape, x, y, width, height);
      cs.fill();
    } else if (hasLine) {
      cs.setStrokingColor(lineColor);
      cs.setLineWidth(lineWidth);
      appendShapePath(cs, shape, x, y, width, height);
      cs.stroke();
    }

    renderShapeText(cs, shape, x, y, width, height, scaleFactor);
  }

  private void appendShapePath(PDPageContentStream cs, XSSFSimpleShape shape,
      float x, float y, float width, float height) throws IOException {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    STShapeType.Enum shapeType =
        (spPr != null && spPr.isSetPrstGeom()) ? spPr.getPrstGeom().getPrst() : null;

    if (STShapeType.PARALLELOGRAM == shapeType) {
      // Excel computes the horizontal slant as adj * height / 100000, not adj * width / 100000.
      // This keeps the slant angle consistent regardless of shape width.
      float offset = (float) (readAdj(spPr, 0.25) * height);
      cs.moveTo(x + offset, y + height); // top-left
      cs.lineTo(x + width, y + height);  // top-right
      cs.lineTo(x + width - offset, y);  // bottom-right
      cs.lineTo(x, y);                   // bottom-left
      cs.closePath();
    } else if (STShapeType.DIAMOND == shapeType) {
      float cx = x + width / 2f;
      float cy = y + height / 2f;
      cs.moveTo(cx, y + height);
      cs.lineTo(x + width, cy);
      cs.lineTo(cx, y);
      cs.lineTo(x, cy);
      cs.closePath();
    } else if (STShapeType.ROUND_RECT == shapeType) {
      float r = (float) (readAdj(spPr, 0.16667) * Math.min(width, height));
      appendRoundRectPath(cs, x, y, width, height, r);
    } else if (STShapeType.ELLIPSE == shapeType) {
      float kappa = 0.5523f;
      float rx = width / 2f;
      float ry = height / 2f;
      float cx = x + rx;
      float cy = y + ry;
      cs.moveTo(cx, cy + ry);
      cs.curveTo(cx + rx * kappa, cy + ry, cx + rx, cy + ry * kappa, cx + rx, cy);
      cs.curveTo(cx + rx, cy - ry * kappa, cx + rx * kappa, cy - ry, cx, cy - ry);
      cs.curveTo(cx - rx * kappa, cy - ry, cx - rx, cy - ry * kappa, cx - rx, cy);
      cs.curveTo(cx - rx, cy + ry * kappa, cx - rx * kappa, cy + ry, cx, cy + ry);
      cs.closePath();
    } else {
      cs.addRect(x, y, width, height);
    }
  }

  private double readAdj(CTShapeProperties spPr, double defaultVal) {
    if (!spPr.isSetPrstGeom()) {
      return defaultVal;
    }
    CTPresetGeometry2D prstGeom = spPr.getPrstGeom();
    if (!prstGeom.isSetAvLst()) {
      return defaultVal;
    }
    for (CTGeomGuide gd : prstGeom.getAvLst().getGdList()) {
      if ("adj".equals(gd.getName()) && gd.getFmla() != null) {
        String fmla = gd.getFmla().trim();
        if (fmla.startsWith("val ")) {
          try {
            return Long.parseLong(fmla.substring(4)) / 100000.0;
          } catch (NumberFormatException e) {
            // fall through to default
          }
        }
      }
    }
    return defaultVal;
  }

  private void appendRoundRectPath(PDPageContentStream cs,
      float x, float y, float width, float height, float r) throws IOException {
    float k = 0.5523f * r;
    cs.moveTo(x + r, y);
    cs.lineTo(x + width - r, y);
    cs.curveTo(x + width - r + k, y, x + width, y + k, x + width, y + r);
    cs.lineTo(x + width, y + height - r);
    cs.curveTo(x + width, y + height - r + k, x + width - r + k, y + height,
        x + width - r, y + height);
    cs.lineTo(x + r, y + height);
    cs.curveTo(x + r - k, y + height, x, y + height - r + k, x, y + height - r);
    cs.lineTo(x, y + r);
    cs.curveTo(x, y + r - k, x + r - k, y, x + r, y);
    cs.closePath();
  }

  private Color getShapeFillColor(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || spPr.isSetNoFill() || !spPr.isSetSolidFill()) {
      return null;
    }
    CTSolidColorFillProperties fill = spPr.getSolidFill();
    if (!fill.isSetSrgbClr()) {
      return null;
    }
    byte[] rgb = fill.getSrgbClr().getVal();
    return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
  }

  private Color getShapeLineColor(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || !spPr.isSetLn()) {
      return null;
    }
    CTLineProperties ln = spPr.getLn();
    if (ln.isSetNoFill() || !ln.isSetSolidFill()) {
      return null;
    }
    CTSolidColorFillProperties fill = ln.getSolidFill();
    if (!fill.isSetSrgbClr()) {
      return null;
    }
    byte[] rgb = fill.getSrgbClr().getVal();
    return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
  }

  private float getShapeLineWidth(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || !spPr.isSetLn()) {
      return 0f;
    }
    CTLineProperties ln = spPr.getLn();
    if (!ln.isSetW()) {
      return 0f;
    }
    return ln.getW() / 12700f; // EMU to points
  }

  private void renderShapeText(PDPageContentStream cs, XSSFSimpleShape shape,
      float x, float y, float width, float height, float scaleFactor) throws IOException {

    java.util.List<XSSFTextParagraph> paragraphs = shape.getTextParagraphs();
    if (paragraphs.isEmpty()) {
      return;
    }

    // Collect CT text body once for inset and RTL lookup
    var txBody = shape.getCTShape().isSetTxBody() ? shape.getCTShape().getTxBody() : null;

    // Read text body insets from bodyPr (DrawingML defaults: lIns=rIns=91440, tIns=bIns=45720 EMU)
    float leftInset = CELL_PADDING;
    float rightInset = CELL_PADDING;
    float topInset = CELL_PADDING;
    var bodyPr = (txBody != null) ? txBody.getBodyPr() : null;
    if (bodyPr != null) {
      long leftIns = 91440L;
      long rightIns = 91440L;
      long topIns = 45720L;
      if (bodyPr.isSetLIns()) {
        leftIns = ((Number) bodyPr.getLIns()).longValue();
      }
      if (bodyPr.isSetRIns()) {
        rightIns = ((Number) bodyPr.getRIns()).longValue();
      }
      if (bodyPr.isSetTIns()) {
        topIns = ((Number) bodyPr.getTIns()).longValue();
      }
      leftInset = leftIns / 12700f * scaleFactor;
      rightInset = rightIns / 12700f * scaleFactor;
      topInset = topIns / 12700f * scaleFactor;
    }

    // Start rendering from inside the top inset of the shape
    float curY = y + height - topInset;

    for (int paraIdx = 0; paraIdx < paragraphs.size(); paraIdx++) {
      XSSFTextParagraph para = paragraphs.get(paraIdx);
      java.util.List<XSSFTextRun> runs = para.getTextRuns();
      if (runs.isEmpty()) {
        continue;
      }

      // Collect text and font properties from runs
      StringBuilder sb = new StringBuilder();
      for (XSSFTextRun run : runs) {
        sb.append(run.getText());
      }
      String lineText = sb.toString();
      if (lineText.isBlank()) {
        continue;
      }

      XSSFTextRun firstRun = runs.get(0);
      double fontSizePt = firstRun.getFontSize();
      if (fontSizePt <= 0) {
        fontSizePt = 11.0;
      }
      float fontSize = (float) fontSizePt * scaleFactor;
      boolean bold = firstRun.isBold();
      PDType0Font font = fontManager.getFont(bold);

      Color textColor = firstRun.getFontColor();
      if (textColor == null) {
        textColor = Color.BLACK;
      }

      float textWidth;
      try {
        textWidth = font.getStringWidth(lineText) / 1000f * fontSize;
      } catch (Exception e) {
        textWidth = width;
      }

      // In RTL paragraphs (rtl="1"), Excel renders algn="l" as center alignment.
      TextAlign align = para.getTextAlign();
      var ctPara = (txBody != null && paraIdx < txBody.sizeOfPArray())
          ? txBody.getPArray(paraIdx) : null;
      boolean paraRtl = ctPara != null && ctPara.isSetPPr()
          && ctPara.getPPr().isSetRtl() && ctPara.getPPr().getRtl();
      if (paraRtl && align == TextAlign.LEFT) {
        align = TextAlign.CENTER;
      }

      float textX;
      if (align == TextAlign.CENTER) {
        float availWidth = width - leftInset - rightInset;
        textX = x + leftInset + Math.max(0f, (availWidth - textWidth) / 2f);
      } else if (align == TextAlign.RIGHT) {
        textX = x + width - rightInset - textWidth;
      } else {
        textX = x + leftInset;
      }

      final float ascent = font.getFontDescriptor().getAscent() / 1000f * fontSize;
      final float descent = font.getFontDescriptor().getDescent() / 1000f * fontSize;
      curY -= ascent;

      cs.beginText();
      cs.setFont(font, fontSize);
      cs.setNonStrokingColor(textColor);
      cs.newLineAtOffset(textX, curY);
      try {
        cs.showText(lineText);
      } catch (Exception e) {
        // Skip text that cannot be rendered with the current font
      }
      cs.endText();

      curY += descent;
      curY -= CELL_PADDING; // line spacing between paragraphs
    }
  }

  // -------------------------------------------------------------------------
  // Cell rendering
  // -------------------------------------------------------------------------

  private void renderCell(PDPageContentStream cs, Cell cell,
      float x, float y, float width, float height, float scaleFactor) throws IOException {

    CellStyle style = (cell != null) ? cell.getCellStyle() : null;

    // 1. Background fill
    if (style != null) {
      Color bgColor = getBackgroundColor(style);
      if (bgColor != null) {
        cs.setNonStrokingColor(bgColor);
        cs.addRect(x, y, width, height);
        cs.fill();
      }
    }

    // 2. Text
    if (cell != null) {
      String value = getCellDisplayValue(cell);
      if (value != null && !value.isBlank()) {
        renderText(cs, cell, value, x, y, width, height, scaleFactor);
      }
    }

    // 3. Borders (drawn on top)
    if (style != null) {
      renderBorders(cs, style, x, y, width, height);
    }
  }

  // -------------------------------------------------------------------------
  // Cell value
  // -------------------------------------------------------------------------

  /**
   * Returns the display value of a cell.
   *
   * <p>For formula cells, uses the cached result instead of the formula string,
   * since {@link DataFormatter#formatCellValue(Cell)} returns the formula string
   * when no {@code FormulaEvaluator} is provided.</p>
   *
   * <p>For numeric cells with date-like format strings (e.g. {@code yyyy"年"m"月分"}),
   * applies custom date formatting because POI's {@code DateUtil.isCellDateFormatted}
   * may return {@code false} for Japanese date formats.</p>
   *
   * @param cell the cell
   * @return the formatted display value
   */
  private String getCellDisplayValue(Cell cell) {
    CellType effectiveType = (cell.getCellType() == CellType.FORMULA)
        ? cell.getCachedFormulaResultType() : cell.getCellType();

    if (effectiveType == CellType.NUMERIC) {
      String formatString = cell.getCellStyle().getDataFormatString();
      if (isLikelyDateFormat(formatString)) {
        return formatDateValue(cell.getNumericCellValue(), formatString);
      }
      if (cell.getCellType() == CellType.FORMULA) {
        return dataFormatter.formatRawCellContents(
            cell.getNumericCellValue(),
            cell.getCellStyle().getDataFormat(),
            formatString);
      }
      return dataFormatter.formatCellValue(cell);
    }

    if (cell.getCellType() == CellType.FORMULA) {
      return switch (cell.getCachedFormulaResultType()) {
        case STRING -> cell.getRichStringCellValue().getString();
        case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
        default -> "";
      };
    }

    return dataFormatter.formatCellValue(cell);
  }

  /**
   * Returns {@code true} if the format string looks like a date format.
   *
   * <p>POI's {@code DateUtil.isCellDateFormatted} can return {@code false} for
   * Japanese date format strings such as {@code yyyy"年"m"月分"}.
   * This method checks for year tokens ({@code y}/{@code Y}) outside of quoted
   * literals as a more lenient detection heuristic.</p>
   *
   * @param formatString Excel format string
   * @return {@code true} if the format looks like a date format
   */
  private boolean isLikelyDateFormat(String formatString) {
    if (formatString == null || formatString.isEmpty()) {
      return false;
    }
    // Use only the main section (before the first ;)
    String mainSection = formatString.contains(";")
        ? formatString.substring(0, formatString.indexOf(';')) : formatString;
    // Remove quoted literals like "年" "月分"
    String stripped = mainSection.replaceAll("\"[^\"]*\"", "");
    // Remove locale/color prefixes like [$-411] or [red]
    stripped = stripped.replaceAll("\\[[^\\]]*\\]", "");
    // Year tokens (y/Y) are unambiguous date indicators
    return stripped.contains("y") || stripped.contains("Y");
  }

  /**
   * Formats an Excel date serial number using the given Excel format string.
   *
   * <p>Processes the format string token by token, handling quoted literals and
   * the date tokens {@code y}/{@code Y} (year), {@code m}/{@code M} (month),
   * and {@code d}/{@code D} (day).</p>
   *
   * @param numericValue Excel date serial number
   * @param formatString Excel format string
   * @return formatted date string
   */
  private String formatDateValue(double numericValue, String formatString) {
    LocalDate date = DateUtil.getLocalDateTime(numericValue, false).toLocalDate();
    // Use only the main section (before the first ;)
    String fmt = formatString.contains(";")
        ? formatString.substring(0, formatString.indexOf(';')) : formatString;

    StringBuilder result = new StringBuilder();
    int i = 0;
    while (i < fmt.length()) {
      char c = fmt.charAt(i);
      if (c == '"') {
        // Quoted literal: append content as-is, without the surrounding quotes
        int end = fmt.indexOf('"', i + 1);
        if (end > i) {
          result.append(fmt, i + 1, end);
          i = end + 1;
        } else {
          i++;
        }
      } else if (c == '[') {
        // Locale or color prefix: skip the entire bracket group
        int end = fmt.indexOf(']', i);
        i = (end > i) ? end + 1 : i + 1;
      } else if (c == 'y' || c == 'Y') {
        int count = countConsecutive(fmt, i, c);
        result.append(count >= 4
            ? String.format("%04d", date.getYear())
            : String.format("%02d", date.getYear() % 100));
        i += count;
      } else if (c == 'm' || c == 'M') {
        int count = countConsecutive(fmt, i, c);
        result.append(count >= 2
            ? String.format("%02d", date.getMonthValue())
            : String.valueOf(date.getMonthValue()));
        i += count;
      } else if (c == 'd' || c == 'D') {
        int count = countConsecutive(fmt, i, c);
        result.append(count >= 2
            ? String.format("%02d", date.getDayOfMonth())
            : String.valueOf(date.getDayOfMonth()));
        i += count;
      } else {
        result.append(c);
        i++;
      }
    }
    return result.toString();
  }

  private int countConsecutive(String s, int start, char target) {
    char lower = Character.toLowerCase(target);
    int count = 0;
    while (start + count < s.length()
        && Character.toLowerCase(s.charAt(start + count)) == lower) {
      count++;
    }
    return count;
  }

  // -------------------------------------------------------------------------
  // Background color
  // -------------------------------------------------------------------------

  private Color getBackgroundColor(CellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return toAwtColor(style.getFillForegroundColorColor());
  }

  private Color toAwtColor(org.apache.poi.ss.usermodel.Color poiColor) {
    if (poiColor instanceof XSSFColor xssfColor) {
      // getRGBWithTint() returns the actual displayed color after applying theme tints.
      // Fall back to getRGB() when tint information is unavailable.
      byte[] rgb = xssfColor.getRGBWithTint();
      if (rgb == null) {
        rgb = xssfColor.getRGB();
      }
      if (rgb != null && rgb.length == 3) {
        return new Color(
            Byte.toUnsignedInt(rgb[0]),
            Byte.toUnsignedInt(rgb[1]),
            Byte.toUnsignedInt(rgb[2]));
      }
    }
    return null;
  }

  // -------------------------------------------------------------------------
  // Text rendering
  // -------------------------------------------------------------------------

  private void renderText(PDPageContentStream cs, Cell cell, String value,
      float x, float y, float width, float height, float scaleFactor) throws IOException {

    CellStyle style = cell.getCellStyle();
    Font poiFont = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());

    boolean bold = poiFont.getBold();
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

    float ascent = font.getFontDescriptor().getAscent() / 1000f * fontSize;
    float descent = font.getFontDescriptor().getDescent() / 1000f * fontSize;
    float lineHeight = ascent - descent;

    java.util.List<String> lines;
    if (style.getWrapText()) {
      float maxLineWidth = width - 2 * CELL_PADDING;
      lines = wrapTextToLines(value, font, fontSize, maxLineWidth);
    } else {
      lines = java.util.List.of(value);
    }

    float totalTextHeight = lines.size() * lineHeight;

    // Compute baseline Y of the first line from vertical alignment
    float startY;
    VerticalAlignment vertAlign = style.getVerticalAlignment();
    if (vertAlign == VerticalAlignment.TOP) {
      startY = y + height - CELL_PADDING - ascent;
    } else if (vertAlign == VerticalAlignment.CENTER) {
      startY = y + (height - totalTextHeight) / 2f - descent;
    } else { // BOTTOM
      startY = y + CELL_PADDING - descent + totalTextHeight - lineHeight;
    }

    cs.setFont(font, fontSize);
    cs.setNonStrokingColor(textColor);

    for (String line : lines) {
      // Stop rendering when text exceeds the bottom of the cell
      if (startY + descent < y) {
        break;
      }
      if (line.isEmpty()) {
        startY -= lineHeight;
        continue;
      }
      float textWidth;
      try {
        textWidth = font.getStringWidth(line) / 1000f * fontSize;
      } catch (Exception e) {
        textWidth = width;
      }
      float textX = calculateTextX(style.getAlignment(), cell, x, width, textWidth);
      cs.beginText();
      cs.newLineAtOffset(textX, startY);
      try {
        cs.showText(line);
      } catch (Exception e) {
        // Skip text that cannot be rendered with the current font
      }
      cs.endText();
      startY -= lineHeight;
    }
  }

  /**
   * Breaks a string into lines that fit within {@code maxWidth} points.
   *
   * <p>Handles both explicit {@code \n} line breaks and automatic wrapping. For CJK
   * text (no word spaces), the break is inserted before the character that would
   * cause the line to exceed {@code maxWidth}. For Latin text with spaces, the break
   * tries to fall on the last space that fits.</p>
   *
   * @param text the text to wrap
   * @param font the PDF font used for width calculation
   * @param fontSize the font size in points
   * @param maxWidth the maximum line width in points
   * @return list of wrapped lines
   */
  private java.util.List<String> wrapTextToLines(String text, PDType0Font font,
      float fontSize, float maxWidth) {
    java.util.List<String> result = new ArrayList<>();
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
          int breakAt;
          if (lastSpaceInCurrent > 0) {
            breakAt = lastSpaceInCurrent;
          } else {
            breakAt = current.length() - 1;
          }
          result.add(current.substring(0, breakAt));
          // Restart from the break point, skipping trailing spaces
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

  private float calculateTextX(HorizontalAlignment align, Cell cell,
      float x, float width, float textWidth) {
    // For GENERAL alignment, Excel right-aligns numeric and date values, left-aligns others.
    HorizontalAlignment effective = align;
    if (align == HorizontalAlignment.GENERAL && cell != null) {
      CellType type = cell.getCellType() == CellType.FORMULA
          ? cell.getCachedFormulaResultType() : cell.getCellType();
      if (type == CellType.NUMERIC || type == CellType.BOOLEAN) {
        effective = HorizontalAlignment.RIGHT;
      }
    }
    return switch (effective) {
      case CENTER -> x + (width - textWidth) / 2f;
      case RIGHT -> x + width - textWidth - CELL_PADDING;
      default -> x + CELL_PADDING;
    };
  }

  // -------------------------------------------------------------------------
  // Border rendering
  // -------------------------------------------------------------------------

  private void renderBorders(PDPageContentStream cs, CellStyle style,
      float x, float y, float width, float height) throws IOException {
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
  }

  private Color xssfColorToAwt(XSSFColor xssfColor) {
    if (xssfColor == null) {
      return Color.BLACK;
    }
    Color c = toAwtColor(xssfColor);
    return (c != null) ? c : Color.BLACK;
  }

  private void renderBorderLine(PDPageContentStream cs, BorderStyle borderStyle,
      Color color, float x1, float y1, float x2, float y2) throws IOException {
    if (borderStyle == BorderStyle.NONE) {
      return;
    }
    cs.setStrokingColor(color);
    cs.setLineWidth(getBorderLineWidth(borderStyle));
    cs.moveTo(x1, y1);
    cs.lineTo(x2, y2);
    cs.stroke();
  }

  private float getBorderLineWidth(BorderStyle style) {
    return switch (style) {
      case MEDIUM, MEDIUM_DASHED, MEDIUM_DASH_DOT, MEDIUM_DASH_DOT_DOT -> 1.0f;
      case THICK -> 1.5f;
      default -> 0.5f;
    };
  }
}
