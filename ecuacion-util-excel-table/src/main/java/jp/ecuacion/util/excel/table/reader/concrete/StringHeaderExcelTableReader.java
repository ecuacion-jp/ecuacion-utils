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
package jp.ecuacion.util.excel.table.reader.concrete;

import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.enums.NoDataString;
import jp.ecuacion.util.excel.exception.ExcelAppException;
import jp.ecuacion.util.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.excel.table.reader.ExcelTableReader;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jspecify.annotations.Nullable;

/**
 * Reads tables with one or multiple header rows.
 *
 * <p>This class unifies the former {@code StringOneLineHeaderExcelTableReader} (single header row)
 *     and adds support for multiple header rows with merged-cell handling.</p>
 *
 * <p>The single-row constructor accepts {@code String[]} for backward compatibility.
 *     The multi-row constructor accepts {@code String[][]} where the first index is the
 *     header row (top to bottom) and the second index is the column.</p>
 *
 * <p>All header rows are validated against the expected labels.
 *     Merged cells in the header area are automatically expanded:
 *     each cell in a merged region receives the value of the top-left (master) cell.
 *     A non-merged empty header cell is treated as an error.</p>
 *
 * <p>The last header row determines the column count and is used for
 *     {@link jp.ecuacion.util.excel.table.bean.ExcelColumn} matching.</p>
 */
public class StringHeaderExcelTableReader extends StringExcelTableReader
    implements IfFormatOneLineHeaderExcelTable<String> {

  /** All header rows' labels: {@code headerLabels2d[row][col]}. */
  private String[][] headerLabels2d;

  private NoDataString noDataString;

  /** Set in {@link #read(Workbook)} before delegating to the parent. */
  protected @Nullable Sheet currentSheet;

  /** POI-basis (0-based) start row of the table header, set in {@link #read(Workbook)}. */
  protected int poiBasisHeaderStartRow;

  /**
   * Constructs a new instance with a single header row.
   * The obtained value from an empty cell is {@code null}.
   *
   * @param sheetName sheet name
   * @param headerLabels expected header labels for the single header row
   * @param tableStartRowNumber 1-based row number of the header row, or {@code null} for
   *     auto-detection by the far-left header label
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum number of data rows, or {@code null} for auto-detection
   */
  public StringHeaderExcelTableReader(String sheetName, String[] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize) {
    this(sheetName, new String[][] {headerLabels}, tableStartRowNumber, tableStartColumnNumber,
        tableRowSize, NoDataString.NULL);
  }

  /**
   * Constructs a new instance with a single header row and a specified empty-cell value.
   *
   * @param sheetName sheet name
   * @param headerLabels expected header labels for the single header row
   * @param tableStartRowNumber 1-based row number of the header row, or {@code null} for
   *     auto-detection by the far-left header label
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum number of data rows, or {@code null} for auto-detection
   * @param noDataString the value to return for an empty cell
   */
  public StringHeaderExcelTableReader(String sheetName, String[] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, NoDataString noDataString) {
    this(sheetName, new String[][] {headerLabels}, tableStartRowNumber, tableStartColumnNumber,
        tableRowSize, noDataString);
  }

  /**
   * Constructs a new instance with multiple header rows.
   * The obtained value from an empty cell is {@code null}.
   *
   * <p>All rows in {@code headerLabels} must have the same length,
   *     which becomes the column count of the table.
   *     Consecutive identical values in the same row are treated as a horizontally merged cell
   *     when writing; vertically merged cells (same value in all rows of a column)
   *     can be annotated with a single-element
   *     {@link jp.ecuacion.util.excel.table.bean.ExcelColumn}.</p>
   *
   * @param sheetName sheet name
   * @param headerLabels expected labels for each header row (expanded form):
   *     {@code headerLabels[row][col]}, top row first
   * @param tableStartRowNumber 1-based row number of the first header row,
   *     or {@code null} for auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum number of data rows, or {@code null} for auto-detection
   */
  public StringHeaderExcelTableReader(String sheetName, String[][] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize) {
    this(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        NoDataString.NULL);
  }

  /**
   * Constructs a new instance with multiple header rows and a specified empty-cell value.
   *
   * @param sheetName sheet name
   * @param headerLabels expected labels for each header row (expanded form)
   * @param tableStartRowNumber 1-based row number of the first header row,
   *     or {@code null} for auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum number of data rows, or {@code null} for auto-detection
   * @param noDataString the value to return for an empty cell
   */
  public StringHeaderExcelTableReader(String sheetName, String[][] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, NoDataString noDataString) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, null);

    this.headerLabels2d = ObjectsUtil.requireNonNull(headerLabels);
    this.noDataString = noDataString;

    setTableColumnSize(getHeaderLabels().length);
  }

  /**
   * Returns the last (bottom) header row's labels, matching the column count.
   * This is the row used for {@link jp.ecuacion.util.excel.table.bean.ExcelColumn} matching.
   *
   * @return last header row labels
   */
  @Override
  public String[] getHeaderLabels() {
    return headerLabels2d[headerLabels2d.length - 1];
  }

  /**
   * Returns all header rows' labels as a 2-D array.
   *
   * @return {@code headerLabels[row][col]}, top row first
   */
  public String[][] getHeaderLabels2d() {
    return headerLabels2d;
  }

  @Override
  public String[][] getHeaderLabelData() {
    return headerLabels2d;
  }

  @Override
  public int getNumberOfHeaderLines() {
    return headerLabels2d.length;
  }

  @Override
  public String getFarLeftAndTopHeaderLabel() {
    ObjectsUtil.requireSizeNonZero(headerLabels2d[0]);
    return ObjectsUtil.requireNonNull(headerLabels2d[0][0]);
  }

  @Override
  public NoDataString getNoDataString() {
    return noDataString;
  }

  /**
   * Stores the current sheet and the header start row, then delegates to the parent.
   *
   * <p>Both are used by {@link #validateHeaderData} for merged-cell expansion.</p>
   *
   * @param workbook workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  @Override
  public List<List<String>> read(Workbook workbook) throws EncryptedDocumentException, IOException {
    Sheet sheet = workbook.getSheet(getSheetName());
    if (sheet != null) {
      this.currentSheet = sheet;
      this.poiBasisHeaderStartRow =
          getPoiBasisDeterminedTableStartRowNumber(sheet, tableStartColumnNumber);
    }
    return super.read(workbook);
  }

  /**
   * Expands merged cells in the header area, then validates all header rows.
   *
   * <p>A merged cell is treated as if every cell in its range holds the master cell's value.
   *     After expansion, any remaining {@code null} or empty cell is reported as an error
   *     because it represents an unexpected gap in the header.</p>
   *
   * @param headerData raw header data (modified in place to expand merged cells)
   * @throws ExcelAppException when a header cell is empty without being merged,
   *     or when labels do not match
   */
  @Override
  public void validateHeaderData(List<List<String>> headerData) throws ExcelAppException {
    if (currentSheet != null && headerLabels2d.length > 1) {
      expandMergedCells(headerData);
      checkNoBlankHeaderCells(headerData);
    }
    super.validateHeaderData(headerData);
  }

  /**
   * Removes the header rows from {@code excelData} and returns them.
   *
   * @param excelData all rows including header rows; header rows are removed in place
   * @return the removed header rows
   * @throws ExcelAppException ExcelAppException
   */
  @Override
  public @Nullable List<List<String>> updateAndGetHeaderData(List<List<String>> excelData)
      throws ExcelAppException {
    List<List<String>> headerRows = new ArrayList<>();
    int numHeaderRows = getNumberOfHeaderLines();
    for (int i = 0; i < numHeaderRows && !excelData.isEmpty(); i++) {
      headerRows.add(excelData.remove(0));
    }
    return headerRows;
  }

  /**
   * Expands merged cells in {@code headerData} using the sheet's merged regions.
   *
   * <p>For each merged region that overlaps the header rows of the table,
   *     every cell in the region (within the table's column range) is set to the master
   *     cell's value.</p>
   *
   * @param headerData the raw header data to expand in place
   */
  private void expandMergedCells(List<List<String>> headerData) {
    Sheet sheet = ObjectsUtil.requireNonNull(currentSheet);
    int numHeaderRows = headerLabels2d.length;
    int poiBasisStartCol = tableStartColumnNumber - 1;
    int numCols = getHeaderLabels().length;

    for (CellRangeAddress region : sheet.getMergedRegions()) {
      int regionFirstRow = region.getFirstRow();
      int regionLastRow = region.getLastRow();
      int regionFirstCol = region.getFirstColumn();
      int regionLastCol = region.getLastColumn();

      // Skip regions that don't overlap with the header area.
      if (regionLastRow < poiBasisHeaderStartRow
          || regionFirstRow >= poiBasisHeaderStartRow + numHeaderRows) {
        continue;
      }
      if (regionLastCol < poiBasisStartCol
          || regionFirstCol >= poiBasisStartCol + numCols) {
        continue;
      }

      // Get the master cell value from headerData.
      int masterHeaderRow = regionFirstRow - poiBasisHeaderStartRow;
      int masterHeaderCol = regionFirstCol - poiBasisStartCol;
      if (masterHeaderRow < 0 || masterHeaderRow >= headerData.size()) {
        continue;
      }
      List<String> masterRow = headerData.get(masterHeaderRow);
      if (masterHeaderCol < 0 || masterHeaderCol >= masterRow.size()) {
        continue;
      }
      String masterValue = masterRow.get(masterHeaderCol);

      // Fill the master value into all cells within the header area of this region.
      for (int r = regionFirstRow; r <= regionLastRow; r++) {
        int headerRowIdx = r - poiBasisHeaderStartRow;
        if (headerRowIdx < 0 || headerRowIdx >= headerData.size()) {
          continue;
        }
        List<String> row = headerData.get(headerRowIdx);
        for (int c = regionFirstCol; c <= regionLastCol; c++) {
          int colIdx = c - poiBasisStartCol;
          if (colIdx >= 0 && colIdx < row.size()) {
            row.set(colIdx, masterValue);
          }
        }
      }
    }
  }

  /**
   * Checks that no header cell is null or empty after merged-cell expansion.
   *
   * @param headerData expanded header data
   * @throws ExcelAppException when a blank non-merged cell is found
   */
  private void checkNoBlankHeaderCells(List<List<String>> headerData) throws ExcelAppException {
    for (int rowIdx = 0; rowIdx < headerData.size(); rowIdx++) {
      List<String> row = headerData.get(rowIdx);
      for (int colIdx = 0; colIdx < row.size(); colIdx++) {
        String val = row.get(colIdx);
        if (val == null || val.isEmpty()) {
          int excelRow = poiBasisHeaderStartRow + rowIdx + 1;
          int excelCol = tableStartColumnNumber + colIdx;
          throw new ExcelAppException(
              "jp.ecuacion.util.excel.reader.HeaderCellIsBlank.message",
              getSheetName(), Integer.toString(excelRow), Integer.toString(excelCol));
        }
      }
    }
  }

  /**
   * Returns the column count, using the last header row for multi-row headers.
   *
   * <p>For multi-row headers the first (group) row may contain horizontally merged cells
   *     whose slave cells appear empty.  Using that row to auto-detect the column count
   *     would yield too small a value, so the last header row (individual column labels)
   *     is used instead.</p>
   *
   * @param sheet sheet
   * @param poiBasisDeterminedTableStartRowNumber the 0-based row of the first header row
   * @param poiBasisDeterminedTableStartColumnNumber the 0-based start column
   * @param ignoresColumnSizeSetInReader ignoresColumnSizeSetInReader
   * @return column count
   * @throws ExcelAppException ExcelAppException
   */
  @Override
  public Integer getTableColumnSize(Sheet sheet, int poiBasisDeterminedTableStartRowNumber,
      int poiBasisDeterminedTableStartColumnNumber, boolean ignoresColumnSizeSetInReader)
      throws ExcelAppException {
    if (getNumberOfHeaderLines() > 1) {
      // Merged cells (horizontal and vertical) make slave cells appear empty,
      // so auto-detecting column count from raw cells is unreliable for multi-row headers.
      // The column count is always taken from the constructor-provided headerLabels.
      return getHeaderLabels().length;
    }
    return super.getTableColumnSize(sheet, poiBasisDeterminedTableStartRowNumber,
        poiBasisDeterminedTableStartColumnNumber, ignoresColumnSizeSetInReader);
  }

  @Override
  public StringHeaderExcelTableReader defaultDateTimeFormat(DateTimeFormatter dateTimeFormat) {
    return (StringHeaderExcelTableReader) super.defaultDateTimeFormat(dateTimeFormat);
  }

  @Override
  public StringHeaderExcelTableReader columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringHeaderExcelTableReader) super.columnDateTimeFormat(columnNumber, dateTimeFormat);
  }

  @Override
  public StringHeaderExcelTableReader ignoresAdditionalColumnsOfHeaderData(boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @Override
  public StringHeaderExcelTableReader isVerticalAndHorizontalOpposite(boolean value) {
    this.isVerticalAndHorizontalOpposite = value;
    return this;
  }
}
