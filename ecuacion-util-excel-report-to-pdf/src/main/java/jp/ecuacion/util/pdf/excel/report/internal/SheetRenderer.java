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
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import jp.ecuacion.util.pdf.excel.report.exception.SheetHasNoPrintAreaException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
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

  private static final float PX_TO_PT = 72f / 96f;

  private final PDDocument document;
  private final CellRenderer cellRenderer;
  private final ShapeRenderer shapeRenderer;
  private final HeaderFooterRenderer headerFooterRenderer;
  /**
   * Maximum Digit Width in pixels at 96 DPI, used to convert column char-unit widths to points.
   * 0 means "use POI's getDefaultCharWidth() at render time" (backward-compatible default).
   */
  private final int mdw;

  /**
   * Constructs a {@code SheetRenderer} using POI's built-in MDW calculation.
   *
   * @param document the target PDF document
   * @param fontManager the font manager providing embedded fonts
   * @param sourcePath the source Excel file path, used for {@code &F} and {@code &Z} codes
   * @param dateLocale the locale used to resolve locale-sensitive built-in date formats
   */
  public SheetRenderer(PDDocument document, FontManager fontManager, @Nullable Path sourcePath,
      Locale dateLocale) {
    this(document, fontManager, sourcePath, dateLocale, 0);
  }

  /**
   * Constructs a {@code SheetRenderer} with an explicitly supplied MDW.
   *
   * <p>When {@code mdw > 0} the supplied value is used directly for all column-width
   * calculations instead of POI's {@code getDefaultCharWidth()} (which can underestimate
   * by one pixel for some fonts). Pass {@code 0} to retain the backward-compatible
   * POI-based calculation.</p>
   *
   * @param document the target PDF document
   * @param fontManager the font manager providing embedded fonts
   * @param sourcePath the source Excel file path, used for {@code &F} and {@code &Z} codes
   * @param dateLocale the locale used to resolve locale-sensitive built-in date formats
   * @param mdw maximum digit width in pixels at 96 DPI (0 = use POI default)
   */
  public SheetRenderer(PDDocument document, FontManager fontManager, @Nullable Path sourcePath,
      Locale dateLocale, int mdw) {
    this.document = document;
    this.mdw = mdw;
    CellValueFormatter formatter =
        new CellValueFormatter(new DataFormatter(Locale.US), dateLocale);
    this.cellRenderer = new CellRenderer(fontManager, formatter);
    this.shapeRenderer = new ShapeRenderer(document, fontManager);
    this.headerFooterRenderer = new HeaderFooterRenderer(fontManager, sourcePath);
  }

  /**
   * Renders the specified sheet into the PDF document.
   *
   * @param workbook the workbook containing the sheet
   * @param sheetIndex 0-based index of the sheet
   * @throws IOException if a PDF I/O error occurs
   * @throws PdfGenerateException if the sheet cannot be rendered
   */
  public void render(Workbook workbook, int sheetIndex) throws IOException {
    Sheet sheet = workbook.getSheetAt(sheetIndex);
    cellRenderer.currentWorkbook = (workbook instanceof XSSFWorkbook xssfWb) ? xssfWb : null;

    final List<TableRenderInfo> tableInfos = cellRenderer.collectTableRenderInfos(sheet, workbook);

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

    // Compute natural (unscaled) column widths once — reused by scale computation and rendering.
    float[] naturalColWidths = new float[lastCol - firstCol + 1];
    float naturalColTotal = 0f;
    for (int c = firstCol; c <= lastCol; c++) {
      float w = getColumnNaturalWidthInPt(sheet, c);
      naturalColWidths[c - firstCol] = w;
      naturalColTotal += w;
    }

    float[] naturalRowHeights = new float[lastRow - firstRow + 1];
    float naturalRowTotal = 0f;
    for (int r = firstRow; r <= lastRow; r++) {
      Row row = sheet.getRow(r);
      float h = (row != null) ? row.getHeightInPoints() : sheet.getDefaultRowHeightInPoints();
      naturalRowHeights[r - firstRow] = h;
      naturalRowTotal += h;
    }

    float scaleFactor = computeScaleFactor(sheet, naturalColTotal, naturalRowTotal,
        printableWidth, printableHeight);

    float[] colWidths = new float[naturalColWidths.length];
    for (int i = 0; i < naturalColWidths.length; i++) {
      colWidths[i] = naturalColWidths[i] * scaleFactor;
    }

    // Horizontal centering: when printOptions/@horizontalCentered is set and the scaled content
    // is narrower than the printable width, shift the content right to center it between margins.
    float totalColWidth = 0f;
    for (float w : colWidths) {
      totalColWidth += w;
    }
    boolean horizontalCentered = false;
    if (sheet instanceof XSSFSheet xssfSheet) {
      var printOpts = xssfSheet.getCTWorksheet().getPrintOptions();
      if (printOpts != null) {
        horizontalCentered = printOpts.getHorizontalCentered();
      }
    }
    float centeringOffset =
        (horizontalCentered && totalColWidth < printableWidth)
            ? (printableWidth - totalColWidth) / 2f : 0f;
    final float contentLeftMargin = leftMargin + centeringOffset;

    float[] rowHeights = new float[naturalRowHeights.length];
    for (int i = 0; i < naturalRowHeights.length; i++) {
      rowHeights[i] = naturalRowHeights[i] * scaleFactor;
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
    // Detect preamble rows: rows before the print title row that appear only on page 1.
    // When Print_Titles is set to a row that is not the first row of the print area, the rows
    // preceding the title row (the "preamble") must be rendered on the first page before the
    // repeated header, while subsequent pages show only the title row followed by content.
    int preambleLastRow = (repeatFirst > firstRow) ? repeatFirst - 1 : -1;
    float preambleHeight = 0f;
    if (preambleLastRow >= 0) {
      for (int r = firstRow; r <= preambleLastRow; r++) {
        preambleHeight += rowHeights[r - firstRow];
      }
    }

    // Content rows are those that are NOT title rows and NOT preamble rows.
    int contentFirstRow = (repeatFirst >= 0) ? repeatLast + 1 : firstRow;
    float contentPageHeight = printableHeight - repeatingRowsHeight;

    List<int[]> rowPages;
    if (repeatFirst >= 0 && contentFirstRow <= lastRow && contentPageHeight > 0) {
      float[] contentRowHeights =
          Arrays.copyOfRange(rowHeights, contentFirstRow - firstRow, rowHeights.length);
      float page1ContentHeight = contentPageHeight - preambleHeight;
      rowPages = buildPageRanges(contentFirstRow, lastRow, sheet.getRowBreaks(), contentRowHeights,
          page1ContentHeight > 0 ? page1ContentHeight : contentPageHeight, contentPageHeight);
    } else {
      repeatFirst = -1;
      rowPages = buildPageRanges(firstRow, lastRow, sheet.getRowBreaks(), rowHeights,
          printableHeight, printableHeight);
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

    // Determine whether fitToPage=true is active and fitToWidth constrains to 1 page wide.
    // When fitToPage=true and fitToWidth=1 (the most common case), Excel forces all columns
    // onto a single horizontal page — the fit scale ensures they fit. We must not create
    // automatic column-page breaks in this mode; only explicit manual column breaks apply.
    boolean fitToOnePage = false;
    if (sheet instanceof XSSFSheet xssfSheetFtp) {
      if (xssfSheetFtp.getFitToPage()) {
        long fitToWidth = 1; // OOXML default when fitToPage is active
        if (xssfSheetFtp.getCTWorksheet().isSetPageSetup()) {
          var ctPs = xssfSheetFtp.getCTWorksheet().getPageSetup();
          if (ctPs.isSetFitToWidth()) {
            fitToWidth = ctPs.getFitToWidth();
          }
        }
        fitToOnePage = (fitToWidth == 1);
      }
    }

    // When no manual column breaks are defined and the natural (unscaled) column total
    // fits within the printable width, treat the sheet as single-column-page.
    List<int[]> colPages;
    if (repeatFirstCol >= 0 && contentFirstCol <= lastCol && contentPageWidth > 0) {
      float naturalContentColTotal = naturalColTotal - repeatingColsWidth / scaleFactor;
      float colPageWidth =
          (sheet.getColumnBreaks().length == 0
              && (fitToOnePage || naturalContentColTotal <= contentPageWidth))
              ? Float.MAX_VALUE
              : contentPageWidth;
      float[] contentColWidths =
          Arrays.copyOfRange(colWidths, contentFirstCol - firstCol, colWidths.length);
      colPages = buildPageRanges(contentFirstCol, lastCol, sheet.getColumnBreaks(),
          contentColWidths, colPageWidth, colPageWidth);
    } else {
      repeatFirstCol = -1;
      float colPageWidth =
          (sheet.getColumnBreaks().length == 0
              && (fitToOnePage || naturalColTotal <= printableWidth))
              ? Float.MAX_VALUE
              : printableWidth;
      colPages = buildPageRanges(firstCol, lastCol, sheet.getColumnBreaks(), colWidths,
          colPageWidth, colPageWidth);
    }

    int totalPages = rowPages.size() * colPages.size();
    int pageNumber = 1;
    for (int[] colPage : colPages) {
      for (int[] rowPage : rowPages) {
        renderPage(sheet, pageSize, leftMargin, rightMargin, topMargin,
            headerMarginPt, footerMarginPt, contentLeftMargin, rowPage[0], rowPage[1], colPage[0],
            colPage[1], firstRow, firstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap,
            repeatFirst,
            repeatLast, repeatFirstCol, repeatLastCol, repeatingColsWidth, preambleLastRow,
            pageNumber, totalPages, tableInfos);
        pageNumber++;
      }
    }
  }

  // -------------------------------------------------------------------------
  // Page setup helpers
  // -------------------------------------------------------------------------

  private int[] getPrintAreaBounds(Workbook workbook, Sheet sheet, int sheetIndex) {
    String printArea = workbook.getPrintArea(sheetIndex);

    if (printArea != null && !printArea.isBlank()) {
      String ref =
          printArea.contains("!") ? printArea.substring(printArea.indexOf('!') + 1) : printArea;
      ref = ref.replace("$", "");
      String[] parts = ref.split(":", -1);
      if (parts.length == 2) {
        int firstRow = cellRefToRow(parts[0]);
        int firstCol = cellRefToCol(parts[0]);
        int lastRow = cellRefToRow(parts[1]);
        int lastCol = cellRefToCol(parts[1]);
        return new int[] {firstRow, lastRow, firstCol, lastCol};
      }
    }

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
      throw new SheetHasNoPrintAreaException(sheet.getSheetName());
    }
    return new int[] {firstRow, lastRow, firstCol, lastCol};
  }

  /** Returns the 0-based row index from a cell reference like "A1" or "B10". */
  private int cellRefToRow(String ref) {
    int i = ref.length() - 1;
    while (i >= 0 && Character.isDigit(ref.charAt(i))) {
      i--;
    }
    return Integer.parseInt(ref.substring(i + 1)) - 1;
  }

  /** Returns the 0-based column index from a cell reference like "A1" or "AB3". */
  private int cellRefToCol(String ref) {
    int col = 0;
    for (int i = 0; i < ref.length(); i++) {
      char ch = ref.charAt(i);
      if (Character.isLetter(ch)) {
        col = col * 26 + (Character.toUpperCase(ch) - 'A' + 1);
      }
    }
    return col - 1;
  }

  private PDRectangle getPageSize(PrintSetup ps) {
    return switch (ps.getPaperSize()) {
      case PrintSetup.LETTER_PAPERSIZE -> PDRectangle.LETTER;
      case PrintSetup.A5_PAPERSIZE -> PDRectangle.A5;
      default -> PDRectangle.A4;
    };
  }

  /**
   * Computes the scale factor for rendering a sheet.
   *
   * <p>Priority:</p>
   * <ol>
   *   <li>Explicit {@code scale} attribute in pageSetup → use that value directly.</li>
   *   <li>"Fit to page" mode ({@code fitToPage=true} in sheetPr) → compute the minimum scale
   *       that makes both columns and rows fit within the printable area (equivalent to
   *       {@code min(printableWidth/naturalColTotal, printableHeight/naturalRowTotal)}).</li>
   *   <li>No settings → render at 1:1; scale down only if content exceeds the printable area.</li>
   * </ol>
   *
   * @param naturalColTotal sum of natural (unscaled) column widths in points
   * @param naturalRowTotal sum of natural (unscaled) row heights in points
   */
  private float computeScaleFactor(Sheet sheet, float naturalColTotal, float naturalRowTotal,
      float printableWidth, float printableHeight) {
    // "Fit to page" mode (fitToPage=true in sheetPr) takes priority over the scale attribute.
    // Excel keeps the old scale value in pageSetup even after switching to "Fit to" mode,
    // so we must check fitToPage first to avoid applying the stale scale value.
    boolean fitToPage = (sheet instanceof XSSFSheet xssfSheet0) && xssfSheet0.getFitToPage();

    if (!fitToPage) {
      if (sheet instanceof XSSFSheet xssfSheet && xssfSheet.getCTWorksheet().isSetPageSetup()) {
        var ctPageSetup = xssfSheet.getCTWorksheet().getPageSetup();
        if (ctPageSetup.isSetScale()) {
          long s = ctPageSetup.getScale();
          return (s > 0 && s <= 400) ? s / 100f : 1f;
        }
      }
    }

    float fitScale;
    if (fitToPage && naturalColTotal > 0) {
      // "Fit to page" mode.
      long fitToWidth = 1;
      long fitToHeight = 1;
      if (sheet instanceof XSSFSheet xssfSheet2 && xssfSheet2.getCTWorksheet().isSetPageSetup()) {
        var ctPs2 = xssfSheet2.getCTWorksheet().getPageSetup();
        if (ctPs2.isSetFitToWidth()) {
          fitToWidth = ctPs2.getFitToWidth();
        }
        if (ctPs2.isSetFitToHeight()) {
          fitToHeight = ctPs2.getFitToHeight();
        }
      }
      // Use Excel's stored fit scale when available and it does not overflow the
      // printable width.  Excel saves the scale it computed (using its own MDW) to
      // pageSetup/@scale; if our MDW matches Excel's the cached scale is exactly
      // right.  Only recompute dynamically when the cached scale would make our
      // columns overflow the page (MDW mismatch) or when no scale is cached.
      if (sheet instanceof XSSFSheet xssfSheet1 && xssfSheet1.getCTWorksheet().isSetPageSetup()) {
        var ctPs1 = xssfSheet1.getCTWorksheet().getPageSetup();
        if (ctPs1.isSetScale()) {
          long s = ctPs1.getScale();
          if (s > 0 && s <= 400) {
            float cached = s / 100f;
            boolean horizontalOverflow = fitToWidth > 0
                && naturalColTotal * cached > printableWidth * 1.02f;
            if (!horizontalOverflow) {
              return cached;
            }
          }
        }
      }
      // No usable stored scale (or overflow): compute dynamically.
      fitScale = 1.0f;
      if (fitToWidth != 0 && naturalColTotal > 0) {
        fitScale = Math.min(1.0f, printableWidth / naturalColTotal);
      }
      if (fitToHeight != 0 && naturalRowTotal > 0 && naturalRowTotal * fitScale > printableHeight) {
        fitScale = Math.min(fitScale, printableHeight / naturalRowTotal);
      }
    } else {
      // No fitToPage, no explicit scale → match Excel's default: render at 100% scale.
      // Excel does NOT auto-scale content to fit the printable area when no page-scaling
      // setting is configured; content may overflow to the right or span multiple pages.
      fitScale = 1.0f;
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
   */
  private float getColumnNaturalWidthInPt(Sheet sheet, int col) {
    if (sheet instanceof XSSFSheet xssfSheet) {
      var ws = xssfSheet.getCTWorksheet();
      for (var ctCols : ws.getColsArray()) {
        for (CTCol ctCol : ctCols.getColList()) {
          if (ctCol.isSetCustomWidth() && ctCol.getCustomWidth() && ctCol.getMin() <= col + 1
              && col + 1 <= ctCol.getMax()) {
            if (mdw > 0) {
              // OOXML spec §18.3.1.13:
              //   pixel = Truncate(((256 × width + Truncate(128/MDW)) / 256) × MDW)
              // getColumnWidth() returns width in 256ths of a character.
              // roundingCorrection = Truncate(128/MDW) via integer division — intentional.
              int roundingCorrection = 128 / mdw;
              int px = (int) (((sheet.getColumnWidth(col) + roundingCorrection) / 256.0) * mdw);
              return px * PX_TO_PT;
            }
            return sheet.getColumnWidthInPixels(col) * PX_TO_PT;
          }
        }
      }
      if (ws.isSetSheetFormatPr()) {
        double dcw = ws.getSheetFormatPr().getDefaultColWidth();
        if (dcw > 0) {
          if (mdw > 0) {
            // Same spec formula; dcw is in char units so multiply by 256 first.
            int roundingCorrection = 128 / mdw;
            int px = (int) (((dcw * 256 + roundingCorrection) / 256.0) * mdw);
            return px * PX_TO_PT;
          }
          // 7px is the standard MDW for Calibri 11pt at 96 DPI
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
   * Builds page ranges with separate capacity for the first page ({@code maxSizeFirstPage})
   * and all subsequent pages ({@code maxSizeOtherPages}). Used when the first page has less
   * available space because of preamble rows that appear only on it.
   */
  private List<int[]> buildPageRanges(int first, int last, int[] manualBreaks, float[] sizes,
      float maxSizeFirstPage, float maxSizeOtherPages) {
    Set<Integer> breakSet = new HashSet<>();
    for (int b : manualBreaks) {
      breakSet.add(b);
    }

    List<int[]> pages = new ArrayList<>();
    int pageStart = first;
    float currentSize = 0f;
    float maxSize = maxSizeFirstPage;

    for (int i = first; i <= last; i++) {
      float size = sizes[i - first];

      // Automatic break: adding this row/col would exceed the page.
      // A 0.5pt tolerance prevents spurious breaks caused by float accumulation
      // when fit-to-page scaling produces a total that is just barely over the limit.
      if (currentSize + size > maxSize + 0.5f && i > pageStart) {
        pages.add(new int[] {pageStart, i - 1});
        pageStart = i;
        currentSize = size;
        maxSize = maxSizeOtherPages;
      } else {
        currentSize += size;
      }

      if (breakSet.contains(i) && i < last) {
        pages.add(new int[] {pageStart, i});
        pageStart = i + 1;
        currentSize = 0f;
        maxSize = maxSizeOtherPages;
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
      float topMargin, float headerMarginPt, float footerMarginPt, float contentLeftMargin,
      int firstPageRow, int lastPageRow, int firstPageCol, int lastPageCol, int printFirstRow,
      int printFirstCol, float[] rowHeights, float[] colWidths, float scaleFactor,
      Map<String, CellRangeAddress> mergedRegionMap, int repeatFirst, int repeatLast,
      int repeatFirstCol, int repeatLastCol, float repeatingColsWidth, int preambleLastRow,
      int pageNumber, int totalPages, List<TableRenderInfo> tableInfos)
      throws IOException {

    PDPage page = new PDPage(pageSize);
    document.addPage(page);

    try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
      float currentY = pageSize.getHeight() - topMargin;

      // Render preamble rows (rows before the print title row) only on the first page.
      if (preambleLastRow >= 0 && pageNumber == 1) {
        for (int r = printFirstRow; r <= preambleLastRow; r++) {
          if (repeatFirstCol >= 0) {
            renderRowCells(cs, sheet, r, repeatFirstCol, repeatLastCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin, printFirstRow, preambleLastRow, tableInfos);
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin + repeatingColsWidth, printFirstRow, preambleLastRow, tableInfos);
          } else {
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin, printFirstRow, preambleLastRow, tableInfos);
          }
        }
      }

      // Render print title rows at the top of every page.
      if (repeatFirst >= 0) {
        for (int r = repeatFirst; r <= repeatLast; r++) {
          if (repeatFirstCol >= 0) {
            renderRowCells(cs, sheet, r, repeatFirstCol, repeatLastCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin, repeatFirst, repeatLast, tableInfos);
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin + repeatingColsWidth, repeatFirst, repeatLast, tableInfos);
          } else {
            currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
                printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
                contentLeftMargin, repeatFirst, repeatLast, tableInfos);
          }
        }
      }

      for (int r = firstPageRow; r <= lastPageRow; r++) {
        if (repeatFirstCol >= 0) {
          renderRowCells(cs, sheet, r, repeatFirstCol, repeatLastCol, printFirstRow, printFirstCol,
              rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY, contentLeftMargin,
              firstPageRow, lastPageRow, tableInfos);
          currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
              printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
              contentLeftMargin + repeatingColsWidth, firstPageRow, lastPageRow, tableInfos);
        } else {
          currentY = renderRowCells(cs, sheet, r, firstPageCol, lastPageCol, printFirstRow,
              printFirstCol, rowHeights, colWidths, scaleFactor, mergedRegionMap, currentY,
              contentLeftMargin, firstPageRow, lastPageRow, tableInfos);
        }
      }

      // Shapes are rendered on top of cells.
      // On page 1 with preamble rows, the effective first row of the page is printFirstRow
      // (the start of the preamble); on other pages it is the first content row of that page.
      int shapesEffectiveFirstRow =
          (preambleLastRow >= 0 && pageNumber == 1) ? printFirstRow : firstPageRow;
      shapeRenderer.renderShapes(sheet, cs, pageSize, contentLeftMargin, topMargin,
          shapesEffectiveFirstRow, lastPageRow, firstPageCol, printFirstRow, printFirstCol,
          rowHeights, colWidths, scaleFactor);

      // Header/footer use the actual page margin (not the content centering offset).
      headerFooterRenderer.renderHeaderOrFooter(cs, sheet, pageSize, leftMargin, rightMargin,
          true, headerMarginPt, pageNumber, totalPages);
      headerFooterRenderer.renderHeaderOrFooter(cs, sheet, pageSize, leftMargin, rightMargin,
          false, footerMarginPt, pageNumber, totalPages);
    }
  }

  private float renderRowCells(PDPageContentStream cs, Sheet sheet, int r, int firstPageCol,
      int lastPageCol, int printFirstRow, int printFirstCol, float[] rowHeights, float[] colWidths,
      float scaleFactor, Map<String, CellRangeAddress> mergedRegionMap, float currentY,
      float leftMargin, int pageFirstRow, int pageLastRow, List<TableRenderInfo> tableInfos)
      throws IOException {
    float rowHeight = rowHeights[r - printFirstRow];
    Row row = sheet.getRow(r);

    // Pass 1: backgrounds only — must precede text so overflow text sits on top of backgrounds.
    float currentX = leftMargin;
    for (int c = firstPageCol; c <= lastPageCol; c++) {
      float colWidth = colWidths[c - printFirstCol];
      CellRangeAddress region = mergedRegionMap.get(r + "," + c);
      if (region != null && (region.getFirstRow() != r || region.getFirstColumn() != c)) {
        currentX += colWidth;
        continue;
      }
      float cellWidth = computeCellWidth(region, c, firstPageCol, lastPageCol, colWidths,
          printFirstCol);
      float cellHeight = computeCellHeight(region, pageFirstRow, pageLastRow, rowHeights,
          printFirstRow, rowHeight);
      Cell cell = (row != null) ? row.getCell(c) : null;
      float cellBottomY = currentY - cellHeight;
      TableCellStyle tableStyle = cellRenderer.getTableCellStyle(tableInfos, r, c);
      cellRenderer.renderBackground(cs, cell, currentX, cellBottomY, cellWidth, cellHeight,
          tableStyle);
      currentX += colWidth;
    }

    // Pass 2: text and borders — text may overflow into adjacent empty cells.
    currentX = leftMargin;
    for (int c = firstPageCol; c <= lastPageCol; c++) {
      float colWidth = colWidths[c - printFirstCol];
      CellRangeAddress region = mergedRegionMap.get(r + "," + c);
      if (region != null && (region.getFirstRow() != r || region.getFirstColumn() != c)) {
        currentX += colWidth;
        continue;
      }
      float cellWidth = computeCellWidth(region, c, firstPageCol, lastPageCol, colWidths,
          printFirstCol);
      float cellHeight = computeCellHeight(region, pageFirstRow, pageLastRow, rowHeights,
          printFirstRow, rowHeight);
      Cell cell = (row != null) ? row.getCell(c) : null;
      float cellBottomY = currentY - cellHeight;
      TableCellStyle tableStyle = cellRenderer.getTableCellStyle(tableInfos, r, c);
      float overflowWidth = computeTextOverflowWidth(row, cell, c, lastPageCol,
          printFirstCol, colWidths, mergedRegionMap, cellWidth, region);
      cellRenderer.renderForeground(cs, cell, currentX, cellBottomY, cellWidth, cellHeight,
          scaleFactor, tableStyle, overflowWidth);
      // Merged cell boundary borders.
      if (region != null && region.getLastColumn() > region.getFirstColumn()) {
        Cell rightBoundary = (row != null) ? row.getCell(region.getLastColumn()) : null;
        if (rightBoundary != null) {
          CellStyle rightStyle = rightBoundary.getCellStyle();
          XSSFCellStyle rxssf = (rightStyle instanceof XSSFCellStyle s) ? s : null;
          cellRenderer.renderBorderLine(cs, rightStyle.getBorderRight(),
              cellRenderer.xssfColorToAwt(rxssf != null ? rxssf.getRightBorderXSSFColor() : null),
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
          cellRenderer.renderBorderLine(cs, bottomStyle.getBorderBottom(),
              cellRenderer.xssfColorToAwt(bxssf != null ? bxssf.getBottomBorderXSSFColor() : null),
              currentX, cellBottomY, currentX + cellWidth, cellBottomY);
        }
      }
      currentX += colWidth;
    }
    return currentY - rowHeight;
  }

  private float computeCellWidth(@Nullable CellRangeAddress region, int c, int firstPageCol,
      int lastPageCol,
      float[] colWidths, int printFirstCol) {
    if (region == null) {
      return colWidths[c - printFirstCol];
    }
    float w = 0f;
    for (int mc = region.getFirstColumn(); mc <= region.getLastColumn(); mc++) {
      if (mc >= firstPageCol && mc <= lastPageCol) {
        w += colWidths[mc - printFirstCol];
      }
    }
    return w;
  }

  private float computeCellHeight(@Nullable CellRangeAddress region, int pageFirstRow,
      int pageLastRow, float[] rowHeights, int printFirstRow, float defaultHeight) {
    if (region == null) {
      return defaultHeight;
    }
    float h = 0f;
    for (int mr = region.getFirstRow(); mr <= region.getLastRow(); mr++) {
      if (mr >= pageFirstRow && mr <= pageLastRow) {
        h += rowHeights[mr - printFirstRow];
      }
    }
    return h;
  }

  /**
   * Computes the width available for text rendering, including overflow into adjacent empty cells.
   *
   * <p>Excel allows text from a cell to visually overflow into adjacent cells to the right
   * as long as those cells have no user-visible content (value, formula, or non-blank string).
   * Overflow stops at the first non-empty cell, a merged-cell boundary, or the page edge.</p>
   */
  private float computeTextOverflowWidth(@Nullable Row row, @Nullable Cell cell, int col,
      int lastPageCol, int printFirstCol, float[] colWidths,
      Map<String, CellRangeAddress> mergedRegionMap, float cellWidth,
      @Nullable CellRangeAddress cellRegion) {
    // Only non-empty text cells can overflow. Merged cells use their full merged width as-is.
    if (cell == null || cellRegion != null) {
      return cellWidth;
    }
    CellStyle style = cell.getCellStyle();
    if (style.getWrapText() || style.getShrinkToFit()) {
      return cellWidth;
    }
    // Only LEFT / GENERAL alignment overflows to the right.
    HorizontalAlignment align = cellRenderer.getHorizontalAlignment(cell, style);
    // GENERAL-aligned numeric/boolean cells are effectively right-aligned (matching
    // calculateTextX behaviour), so they must not overflow into adjacent empty cells.
    if (align == HorizontalAlignment.GENERAL) {
      CellType type = cell.getCellType() == CellType.FORMULA ? cell.getCachedFormulaResultType()
          : cell.getCellType();
      if (type == CellType.NUMERIC || type == CellType.BOOLEAN) {
        return cellWidth;
      }
    }
    if (align == HorizontalAlignment.RIGHT || align == HorizontalAlignment.CENTER
        || align == HorizontalAlignment.FILL || align == HorizontalAlignment.JUSTIFY
        || align == HorizontalAlignment.DISTRIBUTED) {
      return cellWidth;
    }

    float totalWidth = cellWidth;
    int rowNum = cell.getRowIndex();
    for (int c = col + 1; c <= lastPageCol; c++) {
      if (mergedRegionMap.containsKey(rowNum + "," + c)) {
        break; // merged cell area — stop overflow
      }
      Cell adjCell = (row != null) ? row.getCell(c) : null;
      if (!isCellContentEmpty(adjCell)) {
        break; // adjacent cell has content — stop overflow
      }
      totalWidth += colWidths[c - printFirstCol];
    }
    return totalWidth;
  }

  private static boolean isCellContentEmpty(@Nullable Cell cell) {
    if (cell == null) {
      return true;
    }
    return switch (cell.getCellType()) {
      case BLANK -> true;
      case STRING -> cell.getStringCellValue().isBlank();
      default -> false; // NUMERIC, BOOLEAN, FORMULA, ERROR → has content
    };
  }
}
