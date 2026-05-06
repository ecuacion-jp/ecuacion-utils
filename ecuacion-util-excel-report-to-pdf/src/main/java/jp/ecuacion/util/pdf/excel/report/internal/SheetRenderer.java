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
import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import javax.imageio.ImageIO;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.util.Matrix;
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
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBorder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/**
 * Renders an Excel sheet into one or more PDF pages in a {@link PDDocument}.
 *
 * <p>Rendering respects the sheet's print area, page setup (paper size, orientation, margins,
 * scale), manual page breaks, cell background colors, borders (including dashed styles and
 * diagonal lines), text styles, and embedded pictures (PNG, JPEG, and other common
 * formats supported by {@link javax.imageio.ImageIO}).</p>
 */
public class SheetRenderer {

  /** Pixels-to-points conversion factor (96 DPI screen to 72 DPI points). */
  private static final float PX_TO_PT = 72f / 96f;

  /** Start date of the Reiwa era. Dates before this are not supported for era formatting. */
  private static final LocalDate REIWA_START = LocalDate.of(2019, 5, 1);

  /** Horizontal padding inside a cell, in points. */
  private static final float CELL_PADDING = 2f;

  private final PDDocument document;
  private final FontManager fontManager;
  private final DataFormatter dataFormatter = new DataFormatter(Locale.US);
  private final @Nullable Path sourcePath;

  /** Represents a single formatted text run in a header or footer section. */
  private record HfRun(String text, boolean bold, float fontSize, Color color, boolean underline,
      boolean doubleUnderline, boolean strikethrough, boolean superscript, boolean subscript) {
  }

  /**
   * Constructs a {@code SheetRenderer}.
   *
   * @param document the target PDF document
   * @param fontManager the font manager providing embedded fonts
   * @param sourcePath the source Excel file path, used for {@code &F} and {@code &Z} codes
   */
  public SheetRenderer(PDDocument document, FontManager fontManager, @Nullable Path sourcePath) {
    this.document = document;
    this.fontManager = fontManager;
    this.sourcePath = sourcePath;
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
    PDRectangle pageSize =
        ps.getLandscape() ? new PDRectangle(basePageSize.getHeight(), basePageSize.getWidth())
            : basePageSize;

    float leftMargin = (float) (sheet.getMargin(PageMargin.LEFT) * 72);
    float rightMargin = (float) (sheet.getMargin(PageMargin.RIGHT) * 72);
    float topMargin = (float) (sheet.getMargin(PageMargin.TOP) * 72);
    float bottomMargin = (float) (sheet.getMargin(PageMargin.BOTTOM) * 72);
    final float headerMarginPt = (float) (sheet.getMargin(PageMargin.HEADER) * 72);
    final float footerMarginPt = (float) (sheet.getMargin(PageMargin.FOOTER) * 72);
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

    final Map<String, CellRangeAddress> mergedRegionMap = buildMergedRegionMap(sheet);

    // Detect print title rows (rows that repeat at the top of every page).
    CellRangeAddress repeatingRowsRef = sheet.getRepeatingRows();
    int repeatFirst = -1;
    int repeatLast = -1;
    float repeatingRowsHeight = 0f;
    if (repeatingRowsRef != null) {
      int rf = Math.max(repeatingRowsRef.getFirstRow(), firstRow);
      int rl = Math.min(repeatingRowsRef.getLastRow(), lastRow);
      if (rf <= rl) {
        repeatFirst = rf;
        repeatLast = rl;
        for (int r = repeatFirst; r <= repeatLast; r++) {
          repeatingRowsHeight += rowHeights[r - firstRow];
        }
      }
    }
    // Content rows are those that are NOT title rows.
    // Each page can hold (printableHeight − repeatingRowsHeight) of content rows.
    int contentFirstRow = (repeatFirst >= 0) ? repeatLast + 1 : firstRow;
    float contentPageHeight = printableHeight - repeatingRowsHeight;

    List<int[]> rowPages;
    if (repeatFirst >= 0 && contentFirstRow <= lastRow && contentPageHeight > 0) {
      rowPages = buildPageRanges(contentFirstRow, lastRow, sheet.getRowBreaks(), rowHeights,
          contentPageHeight);
    } else {
      repeatFirst = -1; // disable; fall back to normal pagination
      rowPages =
          buildPageRanges(firstRow, lastRow, sheet.getRowBreaks(), rowHeights, printableHeight);
    }
    // Detect print title columns (columns that repeat at the left of every page).
    CellRangeAddress repeatingColsRef = sheet.getRepeatingColumns();
    int repeatFirstCol = -1;
    int repeatLastCol = -1;
    float repeatingColsWidth = 0f;
    if (repeatingColsRef != null) {
      int cf = Math.max(repeatingColsRef.getFirstColumn(), firstCol);
      int cl = Math.min(repeatingColsRef.getLastColumn(), lastCol);
      if (cf <= cl) {
        repeatFirstCol = cf;
        repeatLastCol = cl;
        for (int c = repeatFirstCol; c <= repeatLastCol; c++) {
          repeatingColsWidth += colWidths[c - firstCol];
        }
      }
    }
    int contentFirstCol = (repeatFirstCol >= 0) ? repeatLastCol + 1 : firstCol;
    float contentPageWidth = printableWidth - repeatingColsWidth;

    // When no manual column breaks are defined and the natural (unscaled) column total
    // fits within the printable width, treat the sheet as single-column-page. An explicit
    // scale > 1 may push the scaled total slightly over the boundary, but the intent is
    // to print all columns on one page — consistent with Excel's behavior.
    List<int[]> colPages;
    if (repeatFirstCol >= 0 && contentFirstCol <= lastCol && contentPageWidth > 0) {
      float naturalContentColTotal = naturalColTotal - repeatingColsWidth / scaleFactor;
      float colPageWidth =
          (sheet.getColumnBreaks().length == 0 && naturalContentColTotal <= contentPageWidth)
              ? Float.MAX_VALUE
              : contentPageWidth;
      colPages = buildPageRanges(contentFirstCol, lastCol, sheet.getColumnBreaks(), colWidths,
          colPageWidth);
    } else {
      repeatFirstCol = -1;
      float colPageWidth =
          (sheet.getColumnBreaks().length == 0 && naturalColTotal <= printableWidth)
              ? Float.MAX_VALUE
              : printableWidth;
      colPages =
          buildPageRanges(firstCol, lastCol, sheet.getColumnBreaks(), colWidths, colPageWidth);
    }

    int totalPages = rowPages.size() * colPages.size();
    int pageNumber = 1;
    for (int[] colPage : colPages) {
      for (int[] rowPage : rowPages) {
        renderPage(sheet, pageSize, leftMargin, rightMargin, topMargin, bottomMargin,
            headerMarginPt, footerMarginPt, rowPage[0], rowPage[1], colPage[0], colPage[1],
            firstRow, firstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, repeatFirst,
            repeatLast, repeatFirstCol, repeatLastCol, repeatingColsWidth, pageNumber, totalPages);
        pageNumber++;
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
      String ref =
          printArea.contains("!") ? printArea.substring(printArea.indexOf('!') + 1) : printArea;
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
  private float computeScaleFactor(Sheet sheet, PrintSetup ps, int firstRow, int lastRow,
      int firstCol, int lastCol, float printableWidth, float printableHeight) {
    // For XSSF sheets, check whether the scale attribute is explicitly present in the XML.
    if (sheet instanceof XSSFSheet xssfSheet && xssfSheet.getCTWorksheet().isSetPageSetup()) {
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
          if (ctCol.isSetCustomWidth() && ctCol.getCustomWidth() && ctCol.getMin() <= col + 1
              && col + 1 <= ctCol.getMax()) {
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
  private List<int[]> buildPageRanges(int first, int last, int[] manualBreaks, float[] sizes,
      float maxSize) {
    Set<Integer> breakSet = new HashSet<>();
    for (int b : manualBreaks) {
      breakSet.add(b);
    }

    List<int[]> pages = new ArrayList<>();
    int pageStart = first;
    float currentSize = 0f;

    for (int i = first; i <= last; i++) {
      float size = sizes[i - first];

      // Automatic break: adding this row/col would exceed the page.
      // A 0.5pt tolerance prevents spurious breaks caused by float accumulation
      // when fit-to-page scaling produces a total that is just barely over the limit.
      if (currentSize + size > maxSize + 0.5f && i > pageStart) {
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

  private void renderPage(Sheet sheet, PDRectangle pageSize, float leftMargin, float rightMargin,
      float topMargin, float bottomMargin, float headerMarginPt, float footerMarginPt,
      int firstPageRow, int lastPageRow, int firstPageCol, int lastPageCol, int printFirstRow,
      int printFirstCol, float[] rowHeights, float[] colWidths, float scaleFactor,
      Map<String, CellRangeAddress> mergedRegionMap, int repeatFirst, int repeatLast,
      int repeatFirstCol, int repeatLastCol, float repeatingColsWidth, int pageNumber,
      int totalPages) throws IOException {

    PDPage page = new PDPage(pageSize);
    document.addPage(page);

    try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
      float currentY = pageSize.getHeight() - topMargin;

      // Render print title rows at the top of every page.
      if (repeatFirst >= 0) {
        for (int r = repeatFirst; r <= repeatLast; r++) {
          if (repeatFirstCol >= 0) {
            renderRowCells(cs, sheet, r, repeatFirstCol, repeatLastCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                leftMargin, repeatFirst, repeatLast);
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                leftMargin + repeatingColsWidth, repeatFirst, repeatLast);
          } else {
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                leftMargin, repeatFirst, repeatLast);
          }
        }
      }

      for (int r = firstPageRow; r <= lastPageRow; r++) {
        if (repeatFirstCol >= 0) {
          renderRowCells(cs, sheet, r, repeatFirstCol, repeatLastCol, printFirstRow, printFirstCol,
              rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY, leftMargin,
              firstPageRow, lastPageRow);
          currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
              printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
              leftMargin + repeatingColsWidth, firstPageRow, lastPageRow);
        } else {
          currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
              printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
              leftMargin, firstPageRow, lastPageRow);
        }
      }

      // Shapes are rendered on top of cells
      renderShapes(sheet, cs, pageSize, leftMargin, topMargin, firstPageRow, lastPageRow,
          firstPageCol, lastPageCol, printFirstRow, printFirstCol, rowHeights, colWidths,
          scaleFactor);

      // Header and footer
      renderHeaderOrFooter(cs, sheet, pageSize, leftMargin, rightMargin, true, headerMarginPt,
          pageNumber, totalPages);
      renderHeaderOrFooter(cs, sheet, pageSize, leftMargin, rightMargin, false, footerMarginPt,
          pageNumber, totalPages);
    }
  }

  private float renderRowCells(PDPageContentStream cs, Sheet sheet, int r, int firstPageCol,
      int lastPageCol, int printFirstRow, int printFirstCol, float[] rowHeights, float[] colWidths,
      float scaleFactor, Map<String, CellRangeAddress> mergedRegionMap, float currentY,
      float leftMargin, int pageFirstRow, int pageLastRow) throws IOException {
    float rowHeight = rowHeights[r - printFirstRow];
    float currentX = leftMargin;
    Row row = sheet.getRow(r);

    for (int c = firstPageCol; c <= lastPageCol; c++) {
      float colWidth = colWidths[c - printFirstCol];
      CellRangeAddress region = mergedRegionMap.get(r + "," + c);
      if (region != null && (region.getFirstRow() != r || region.getFirstColumn() != c)) {
        currentX += colWidth;
        continue;
      }
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
          if (mr >= pageFirstRow && mr <= pageLastRow) {
            cellHeight += rowHeights[mr - printFirstRow];
          }
        }
      }
      Cell cell = (row != null) ? row.getCell(c) : null;
      float cellBottomY = currentY - cellHeight;
      renderCell(cs, cell, currentX, cellBottomY, cellWidth, cellHeight, scaleFactor);
      // For merged cells, right/bottom borders come from the boundary cells.
      if (region != null && region.getLastColumn() > region.getFirstColumn()) {
        Cell rightBoundary = (row != null) ? row.getCell(region.getLastColumn()) : null;
        if (rightBoundary != null) {
          CellStyle rightStyle = rightBoundary.getCellStyle();
          XSSFCellStyle rxssf = (rightStyle instanceof XSSFCellStyle s) ? s : null;
          renderBorderLine(cs, rightStyle.getBorderRight(),
              xssfColorToAwt(rxssf != null ? rxssf.getRightBorderXSSFColor() : null),
              currentX + cellWidth, cellBottomY, currentX + cellWidth, cellBottomY + cellHeight);
        }
      }
      if (region != null && region.getLastRow() > region.getFirstRow()) {
        Row lastMergeRow = sheet.getRow(region.getLastRow());
        Cell bottomBoundary =
            (lastMergeRow != null) ? lastMergeRow.getCell(region.getFirstColumn()) : null;
        if (bottomBoundary != null) {
          CellStyle bottomStyle = bottomBoundary.getCellStyle();
          XSSFCellStyle bxssf = (bottomStyle instanceof XSSFCellStyle s) ? s : null;
          renderBorderLine(cs, bottomStyle.getBorderBottom(),
              xssfColorToAwt(bxssf != null ? bxssf.getBottomBorderXSSFColor() : null), currentX,
              cellBottomY, currentX + cellWidth, cellBottomY);
        }
      }
      currentX += colWidth;
    }
    return currentY - rowHeight;
  }

  // -------------------------------------------------------------------------
  // Shape rendering
  // -------------------------------------------------------------------------

  private void renderShapes(Sheet sheet, PDPageContentStream cs, PDRectangle pageSize,
      float leftMargin, float topMargin, int firstPageRow, int lastPageRow, int firstPageCol,
      int lastPageCol, int printFirstRow, int printFirstCol, float[] rowHeights, float[] colWidths,
      float scaleFactor) throws IOException {

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
      if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor)) {
        continue;
      }

      // Shape bounds in Excel coordinates (from print area start).
      // For two-cell anchors (row2 or col2 > 0) derive the size from the second anchor so
      // that the shape scales correctly with the row/column layout.
      // For one-cell anchors (row2 == col2 == 0) fall back to xfrm.ext or image dimensions.
      float shapeTopY = shapeExcelY(anchor.getRow1(), anchor.getDy1(), rowHeights, printFirstRow);
      float shapeLeftX = shapeExcelX(anchor.getCol1(), anchor.getDx1(), colWidths, printFirstCol);
      float shapeBottomY;
      float shapeRightX;

      if (anchor.getRow2() > 0 || anchor.getCol2() > 0) {
        shapeBottomY = shapeExcelY(anchor.getRow2(), anchor.getDy2(), rowHeights, printFirstRow);
        shapeRightX = shapeExcelX(anchor.getCol2(), anchor.getDx2(), colWidths, printFirstCol);
      } else if (shape instanceof XSSFPicture picture) {
        Dimension dim = picture.getImageDimension();
        if (dim == null || dim.width <= 0 || dim.height <= 0) {
          continue;
        }
        shapeBottomY = shapeTopY + dim.height * PX_TO_PT * scaleFactor;
        shapeRightX = shapeLeftX + dim.width * PX_TO_PT * scaleFactor;
      } else if (shape instanceof XSSFSimpleShape simpleShape) {
        long extCx = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCx();
        long extCy = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCy();
        shapeBottomY = shapeTopY + extCy / 12700f * scaleFactor;
        shapeRightX = shapeLeftX + extCx / 12700f * scaleFactor;
      } else {
        continue;
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

      if (shape instanceof XSSFPicture picture) {
        renderPicture(cs, picture, pdfShapeLeft, pdfShapeBottom, pdfShapeWidth, pdfShapeHeight);
      } else if (shape instanceof XSSFSimpleShape simpleShape) {
        renderShape(cs, simpleShape, pdfShapeLeft, pdfShapeBottom, pdfShapeWidth, pdfShapeHeight,
            scaleFactor);
      }
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

  private void renderPicture(PDPageContentStream cs, XSSFPicture picture, float x, float y,
      float width, float height) throws IOException {
    XSSFPictureData picData = picture.getPictureData();
    byte[] imageBytes = picData.getData();
    String mime = picData.getMimeType();
    PDImageXObject pdImage;
    if ("image/jpeg".equalsIgnoreCase(mime) || "image/jpg".equalsIgnoreCase(mime)) {
      pdImage = JPEGFactory.createFromByteArray(document, imageBytes);
    } else {
      BufferedImage bi = ImageIO.read(new ByteArrayInputStream(imageBytes));
      if (bi == null) {
        return; // ImageIO returned null: unsupported or unreadable format
      }
      pdImage = LosslessFactory.createFromImage(document, bi);
    }
    cs.drawImage(pdImage, x, y, width, height);
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

  private void renderShape(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

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

  private void appendShapePath(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height) throws IOException {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    STShapeType.Enum shapeType =
        (spPr != null && spPr.isSetPrstGeom()) ? spPr.getPrstGeom().getPrst() : null;

    if (STShapeType.PARALLELOGRAM == shapeType) {
      // Excel computes the horizontal slant as adj * height / 100000, not adj * width / 100000.
      // This keeps the slant angle consistent regardless of shape width.
      float offset = (float) (readAdj(spPr, 0.25) * height);
      cs.moveTo(x + offset, y + height); // top-left
      cs.lineTo(x + width, y + height); // top-right
      cs.lineTo(x + width - offset, y); // bottom-right
      cs.lineTo(x, y); // bottom-left
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

  private void appendRoundRectPath(PDPageContentStream cs, float x, float y, float width,
      float height, float r) throws IOException {
    float k = 0.5523f * r;
    cs.moveTo(x + r, y);
    cs.lineTo(x + width - r, y);
    cs.curveTo(x + width - r + k, y, x + width, y + k, x + width, y + r);
    cs.lineTo(x + width, y + height - r);
    cs.curveTo(x + width, y + height - r + k, x + width - r + k, y + height, x + width - r,
        y + height);
    cs.lineTo(x + r, y + height);
    cs.curveTo(x + r - k, y + height, x, y + height - r + k, x, y + height - r);
    cs.lineTo(x, y + r);
    cs.curveTo(x, y + r - k, x + r - k, y, x + r, y);
    cs.closePath();
  }

  private @Nullable Color getShapeFillColor(XSSFSimpleShape shape) {
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

  private @Nullable Color getShapeLineColor(XSSFSimpleShape shape) {
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

  private void renderShapeText(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

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
      var ctPara =
          (txBody != null && paraIdx < txBody.sizeOfPArray()) ? txBody.getPArray(paraIdx) : null;
      boolean paraRtl = ctPara != null && ctPara.isSetPPr() && ctPara.getPPr().isSetRtl()
          && ctPara.getPPr().getRtl();
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

  private void renderCell(PDPageContentStream cs, @Nullable Cell cell, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

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
        if (cell.getCellStyle().getRotation() == 255) {
          renderVerticalText(cs, cell, value, x, y, width, height, scaleFactor);
        } else {
          renderText(cs, cell, value, x, y, width, height, scaleFactor);
        }
      }
    }

    // 3. Borders (drawn on top)
    if (style != null && cell != null) {
      renderBorders(cs, style, cell.getSheet(), x, y, width, height);
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
    CellType effectiveType =
        (cell.getCellType() == CellType.FORMULA) ? cell.getCachedFormulaResultType()
            : cell.getCellType();

    if (effectiveType == CellType.NUMERIC) {
      String formatString = cell.getCellStyle().getDataFormatString();
      if (isLikelyDateFormat(formatString)) {
        return formatDateValue(cell.getNumericCellValue(), formatString);
      }
      if (cell.getCellType() == CellType.FORMULA) {
        return dataFormatter.formatRawCellContents(cell.getNumericCellValue(),
            cell.getCellStyle().getDataFormat(), formatString);
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
    if (formatString == null || formatString.isEmpty()
        || formatString.equalsIgnoreCase("General")) {
      return false;
    }
    // Use only the main section (before the first ;)
    String mainSection =
        formatString.contains(";") ? formatString.substring(0, formatString.indexOf(';'))
            : formatString;
    // Remove quoted literals like "年" "月分"
    String stripped = mainSection.replaceAll("\"[^\"]*\"", "");
    // Remove locale/color prefixes like [$-411] or [red]
    stripped = stripped.replaceAll("\\[[^\\]]*\\]", "");
    // Year tokens (y/Y) and era tokens (g/G) are unambiguous date indicators
    return stripped.contains("y") || stripped.contains("Y") || stripped.contains("g")
        || stripped.contains("G");
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
    var ldt = DateUtil.getLocalDateTime(numericValue, false);
    LocalDate date = ldt.toLocalDate();
    LocalTime time = ldt.toLocalTime();
    Locale locale = extractFormatLocale(formatString);

    String fmt = formatString.contains(";") ? formatString.substring(0, formatString.indexOf(';'))
        : formatString;

    // Pre-scan for AM/PM to enable 12-hour clock mode.
    boolean is12Hour = fmt.toLowerCase(Locale.ENGLISH).contains("am/pm")
        || fmt.toLowerCase(Locale.ENGLISH).contains("a/p");
    int hourVal = is12Hour ? (time.getHour() % 12 == 0 ? 12 : time.getHour() % 12) : time.getHour();

    StringBuilder result = new StringBuilder();
    boolean lastWasHour = false;
    int i = 0;

    while (i < fmt.length()) {
      char c = fmt.charAt(i);
      if (c == '"') {
        int end = fmt.indexOf('"', i + 1);
        if (end > i) {
          result.append(fmt, i + 1, end);
          i = end + 1;
        } else {
          i++;
        }
      } else if (c == '[') {
        int end = fmt.indexOf(']', i);
        i = (end > i) ? end + 1 : i + 1;
      } else if (c == 'y' || c == 'Y') {
        int n = countConsecutive(fmt, i, c);
        result.append(n >= 4 ? String.format("%04d", date.getYear())
            : String.format("%02d", date.getYear() % 100));
        i += n;
        lastWasHour = false;
      } else if (c == 'g' || c == 'G') {
        int n = countConsecutive(fmt, i, c);
        result.append(eraName(date, n));
        i += n;
        lastWasHour = false;
      } else if (c == 'e' || c == 'E') {
        int n = countConsecutive(fmt, i, c);
        result.append(eraYear(date, n));
        i += n;
        lastWasHour = false;
      } else if (c == 'm' || c == 'M') {
        int n = countConsecutive(fmt, i, c);
        if (lastWasHour) {
          result.append(
              n >= 2 ? String.format("%02d", time.getMinute()) : String.valueOf(time.getMinute()));
          lastWasHour = false;
        } else if (n >= 4) {
          result.append(date.getMonth().getDisplayName(TextStyle.FULL, locale));
        } else if (n == 3) {
          result.append(date.getMonth().getDisplayName(TextStyle.SHORT, locale));
        } else {
          result.append(n >= 2 ? String.format("%02d", date.getMonthValue())
              : String.valueOf(date.getMonthValue()));
        }
        i += n;
      } else if (c == 'd' || c == 'D') {
        int n = countConsecutive(fmt, i, c);
        if (n >= 4) {
          result.append(date.getDayOfWeek().getDisplayName(TextStyle.FULL, locale));
        } else if (n == 3) {
          result.append(date.getDayOfWeek().getDisplayName(TextStyle.SHORT, locale));
        } else {
          result.append(n >= 2 ? String.format("%02d", date.getDayOfMonth())
              : String.valueOf(date.getDayOfMonth()));
        }
        i += n;
        lastWasHour = false;
      } else if (c == 'h' || c == 'H') {
        int n = countConsecutive(fmt, i, c);
        result.append(n >= 2 ? String.format("%02d", hourVal) : String.valueOf(hourVal));
        i += n;
        lastWasHour = true;
      } else if (c == 's' || c == 'S') {
        int n = countConsecutive(fmt, i, c);
        result.append(
            n >= 2 ? String.format("%02d", time.getSecond()) : String.valueOf(time.getSecond()));
        i += n;
        lastWasHour = false;
      } else if (c == 'a' || c == 'A') {
        if (i + 4 < fmt.length() && fmt.substring(i, i + 5).equalsIgnoreCase("AM/PM")) {
          result.append(time.getHour() < 12 ? "AM" : "PM");
          i += 5;
        } else if (i + 2 < fmt.length() && fmt.substring(i, i + 3).equalsIgnoreCase("A/P")) {
          result.append(time.getHour() < 12 ? "A" : "P");
          i += 3;
        } else {
          // Japanese weekday: aaa (月) or aaaa (月曜日)
          int n = countConsecutive(fmt, i, c);
          if (n >= 3) {
            int dow = date.getDayOfWeek().getValue() - 1;
            String[] s = {"月", "火", "水", "木", "金", "土", "日"};
            String[] l = {"月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"};
            result.append(n >= 4 ? l[dow] : s[dow]);
            i += n;
          } else {
            result.append(c);
            i++;
          }
        }
      } else {
        result.append(c);
        i++;
      }
    }
    return result.toString();
  }

  private Locale extractFormatLocale(String formatString) {
    int start = formatString.indexOf("[$-");
    if (start >= 0) {
      int end = formatString.indexOf(']', start + 3);
      if (end > start + 3) {
        try {
          int lcid = Integer.parseInt(formatString.substring(start + 3, end), 16) & 0xFFFF;
          if (lcid == 0x0411) {
            return Locale.JAPANESE;
          }
          if (lcid == 0x0407) {
            return Locale.GERMAN;
          }
          if (lcid == 0x040C) {
            return Locale.FRENCH;
          }
          return Locale.ENGLISH;
        } catch (NumberFormatException ignored) {
          // fall through
        }
      }
    }
    return Locale.ENGLISH;
  }

  private String eraName(LocalDate date, int count) {
    if (date.isBefore(REIWA_START)) {
      throw new RuntimeException(
          "Japanese era before Reiwa (2019-05-01) is not supported: " + date);
    }
    return count >= 2 ? "令和" : "令";
  }

  private String eraYear(LocalDate date, int count) {
    if (date.isBefore(REIWA_START)) {
      throw new RuntimeException(
          "Japanese era before Reiwa (2019-05-01) is not supported: " + date);
    }
    int year = date.getYear() - 2018;
    return count >= 2 ? String.format("%02d", year) : String.valueOf(year);
  }

  private int countConsecutive(String s, int start, char target) {
    char lower = Character.toLowerCase(target);
    int count = 0;
    while (start + count < s.length() && Character.toLowerCase(s.charAt(start + count)) == lower) {
      count++;
    }
    return count;
  }

  // -------------------------------------------------------------------------
  // Background color
  // -------------------------------------------------------------------------

  private @Nullable Color getBackgroundColor(CellStyle style) {
    if (style.getFillPattern() != FillPatternType.SOLID_FOREGROUND) {
      return null;
    }
    return toAwtColor(style.getFillForegroundColorColor());
  }

  private @Nullable Color toAwtColor(org.apache.poi.ss.usermodel.Color poiColor) {
    if (poiColor instanceof XSSFColor xssfColor) {
      // getRGBWithTint() returns the actual displayed color after applying theme tints.
      // Fall back to getRGB() when tint information is unavailable.
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

  // -------------------------------------------------------------------------
  // Text rendering
  // -------------------------------------------------------------------------

  private void renderText(PDPageContentStream cs, Cell cell, String value, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

    CellStyle style = cell.getCellStyle();
    Font poiFont = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());

    boolean bold = poiFont.getBold();
    final boolean italic = poiFont.getItalic();
    final boolean strikeout = poiFont.getStrikeout();
    final boolean underline = poiFont.getUnderline() == Font.U_SINGLE
        || poiFont.getUnderline() == Font.U_SINGLE_ACCOUNTING;
    final boolean doubleUnderline = poiFont.getUnderline() == Font.U_DOUBLE
        || poiFont.getUnderline() == Font.U_DOUBLE_ACCOUNTING;
    final boolean accountingUnderline = poiFont.getUnderline() == Font.U_SINGLE_ACCOUNTING
        || poiFont.getUnderline() == Font.U_DOUBLE_ACCOUNTING;
    short typeOffset = poiFont.getTypeOffset();
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

    // Shrink to fit: reduce font size so single-line text fits within the cell width.
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

    // Super/subscript: render at 70% size with a shifted baseline.
    boolean superscript = (typeOffset == Font.SS_SUPER);
    boolean subscript = (typeOffset == Font.SS_SUB);
    float effectiveFontSize = (superscript || subscript) ? fontSize * 0.7f : fontSize;

    float ascent = font.getFontDescriptor().getAscent() / 1000f * effectiveFontSize;
    float descent = font.getFontDescriptor().getDescent() / 1000f * effectiveFontSize;
    float lineHeight = ascent - descent;

    java.util.List<String> lines;
    if (style.getWrapText()) {
      float maxLineWidth = width - 2 * CELL_PADDING;
      lines = wrapTextToLines(value, font, effectiveFontSize, maxLineWidth);
    } else {
      lines = java.util.List.of(value);
    }

    float totalTextHeight = lines.size() * lineHeight;

    float startY;
    VerticalAlignment vertAlign = style.getVerticalAlignment();
    if (vertAlign == VerticalAlignment.TOP) {
      startY = y + height - CELL_PADDING - ascent;
    } else if (vertAlign == VerticalAlignment.CENTER) {
      startY = y + (height - totalTextHeight) / 2f - descent;
    } else { // BOTTOM
      startY = y + CELL_PADDING - descent + totalTextHeight - lineHeight;
    }

    for (String line : lines) {
      // Skip lines above the cell top (BOTTOM/CENTER align overflow upward).
      if (startY > y + height) {
        startY -= lineHeight;
        continue;
      }
      // Stop rendering when text descends below the cell bottom.
      if (startY + descent < y) {
        break;
      }
      if (line.isEmpty()) {
        startY -= lineHeight;
        continue;
      }

      // Shift baseline for super/subscript.
      float lineY = startY;
      if (superscript) {
        lineY += fontSize * 0.35f;
      } else if (subscript) {
        lineY -= fontSize * 0.15f;
      }

      float textWidth;
      try {
        textWidth = font.getStringWidth(line) / 1000f * effectiveFontSize;
      } catch (Exception e) {
        textWidth = width;
      }
      final float textX = calculateTextX(style.getAlignment(), cell, x, width, textWidth);

      cs.beginText();
      cs.setFont(font, effectiveFontSize);
      cs.setNonStrokingColor(textColor);
      if (italic) {
        // Synthetic italic: shear ~12° using tan(12°) ≈ 0.21.
        cs.setTextMatrix(new Matrix(1, 0, 0.21f, 1, textX, lineY));
      } else {
        cs.newLineAtOffset(textX, lineY);
      }
      try {
        cs.showText(line);
      } catch (Exception e) {
        // Skip text that cannot be rendered with the current font.
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
        // Accounting underline spans the full cell width; standard spans the text width.
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
      float width, float height, float scaleFactor) throws IOException {

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
    float centerX = x + width / 2f;
    float currentY = y + height - CELL_PADDING - ascent;

    for (int i = 0; i < value.length(); i++) {
      if (currentY + descent < y) {
        break;
      }
      String ch = String.valueOf(value.charAt(i));
      float charWidth;
      try {
        charWidth = font.getStringWidth(ch) / 1000f * fontSize;
      } catch (Exception e) {
        charWidth = fontSize;
      }
      cs.beginText();
      cs.setFont(font, fontSize);
      cs.setNonStrokingColor(textColor);
      cs.newLineAtOffset(centerX - charWidth / 2f, currentY);
      try {
        cs.showText(ch);
      } catch (Exception e) {
        // Skip characters that cannot be rendered.
      }
      cs.endText();
      currentY -= lineHeight;
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
  private java.util.List<String> wrapTextToLines(String text, PDType0Font font, float fontSize,
      float maxWidth) {
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

  private float calculateTextX(HorizontalAlignment align, Cell cell, float x, float width,
      float textWidth) {
    // For GENERAL alignment, Excel right-aligns numeric and date values, left-aligns others.
    HorizontalAlignment effective = align;
    if (align == HorizontalAlignment.GENERAL && cell != null) {
      CellType type = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType()
          : cell.getCellType();
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

  private void renderBorders(PDPageContentStream cs, CellStyle style, Sheet sheet, float x, float y,
      float width, float height) throws IOException {
    XSSFCellStyle xssfStyle = (style instanceof XSSFCellStyle s) ? s : null;

    renderBorderLine(cs, style.getBorderTop(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getTopBorderXSSFColor() : null), x, y + height,
        x + width, y + height);
    renderBorderLine(cs, style.getBorderBottom(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getBottomBorderXSSFColor() : null), x, y,
        x + width, y);
    renderBorderLine(cs, style.getBorderLeft(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getLeftBorderXSSFColor() : null), x, y, x,
        y + height);
    renderBorderLine(cs, style.getBorderRight(),
        xssfColorToAwt(xssfStyle != null ? xssfStyle.getRightBorderXSSFColor() : null), x + width,
        y, x + width, y + height);
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
    // diagonalDown: top-left → bottom-right (in PDF coords: top = y+height, bottom = y)
    if (ctBorder.getDiagonalDown()) {
      renderBorderLine(cs, diagStyle, color, x, y + height, x + width, y);
    }
    // diagonalUp: bottom-left → top-right (in PDF coords: bottom = y, top = y+height)
    if (ctBorder.getDiagonalUp()) {
      renderBorderLine(cs, diagStyle, color, x, y, x + width, y + height);
    }
  }

  private Color xssfColorToAwt(@Nullable XSSFColor xssfColor) {
    if (xssfColor == null) {
      return Color.BLACK;
    }
    Color c = toAwtColor(xssfColor);
    return (c != null) ? c : Color.BLACK;
  }

  private void renderBorderLine(PDPageContentStream cs, BorderStyle borderStyle, Color color,
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
  // Header / footer rendering
  // -------------------------------------------------------------------------

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
   * @throws IOException if a PDF I/O error occurs
   */
  private void renderHeaderOrFooter(PDPageContentStream cs, Sheet sheet, PDRectangle pageSize,
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

  /**
   * Renders a list of formatted runs starting at {@code startX}.
   *
   * @param cs the content stream
   * @param runs the runs to render
   * @param startX the left edge of the section in points
   * @param baseline the text baseline in PDF coordinates
   * @throws IOException if a PDF I/O error occurs
   */
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
   *
   * @param sectionText the raw section string (after L/C/R splitting)
   * @param pageNumber current 1-based page number
   * @param totalPages total page count
   * @param fileName workbook file name without extension (for {@code &F})
   * @param filePath parent directory path (for {@code &Z})
   * @param sheetName sheet tab name (for {@code &A})
   * @return parsed runs
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
          buf.append(LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy/M/d")));
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          i += 2;
        }
        case 'T' -> {
          flushHfRun(runs, buf, bold, fontSize, color, underline, doubleUnderline, strikethrough,
              superscript, subscript);
          buf.append(LocalTime.now().format(DateTimeFormatter.ofPattern("H:mm")));
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
              // Apply bold/regular from style (font name is ignored; only NotoSansJP available)
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
              // Theme color or other format — skip to next non-code character
              int end = i + 2;
              while (end < sectionText.length()) {
                char ch = sectionText.charAt(end);
                if (!Character.isLetterOrDigit(ch) && ch != '-' && ch != '+') {
                  break;
                }
                end++;
              }
              color = Color.BLACK;
              i = end;
            }
          } else {
            i += 2;
          }
        }
        default -> {
          if (Character.isDigit(code)) {
            // &n... — font size (one or more digits)
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
              // ignore malformed size code
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
