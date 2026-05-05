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
package jp.ecuacion.util.excel.table.writer.concrete;

import java.io.IOException;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.excel.table.reader.concrete.StringHeaderExcelTableReader;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeStringExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with one or multiple header rows.
 *
 * <p>This class unifies the former {@code StringOneLineHeaderExcelTableWriter}
 *     and adds support for multiple header rows.
 *     The single-row constructor accepts {@code String[]} for backward compatibility.</p>
 *
 * <p>When writing, consecutive identical values in the same header row are merged horizontally.
 *     A column whose header label is identical in every header row is merged vertically.</p>
 *
 * <p>The header in the template file is validated against {@code headerLabels} before writing.</p>
 */
public class StringHeaderExcelTableWriter extends ExcelTableWriter<String>
    implements IfDataTypeStringExcelTableWriter, IfFormatOneLineHeaderExcelTable<String> {

  private String[][] headerLabels2d;

  /**
   * Constructs a new instance with a single header row.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels for the single header row
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   *     Must point to the header row of the table.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public StringHeaderExcelTableWriter(String sheetName, String[] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber) {
    this(sheetName, new String[][] {headerLabels}, tableStartRowNumber, tableStartColumnNumber);
  }

  /**
   * Constructs a new instance with multiple header rows.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels: {@code headerLabels[row][col]}, top row first.
   *     All rows must have the same length.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   *     Must point to the first (top) header row of the table.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public StringHeaderExcelTableWriter(String sheetName, String[][] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
    this.headerLabels2d = ObjectsUtil.requireNonNull(headerLabels);
  }

  /**
   * Returns the last header row's labels (single-row backward-compat interface).
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

  /**
   * Validates the template file's headers using a {@link StringHeaderExcelTableReader}.
   *
   * @param workbook workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  @Override
  protected void headerCheck(Workbook workbook) throws EncryptedDocumentException, IOException {
    new StringHeaderExcelTableReader(getSheetName(), headerLabels2d,
        tableStartRowNumber, tableStartColumnNumber, 1)
            .ignoresAdditionalColumnsOfHeaderData(ignoresAdditionalColumnsOfHeaderData())
            .isVerticalAndHorizontalOpposite(isVerticalAndHorizontalOpposite()).read(workbook);
  }

  /**
   * Writes the header rows into the sheet and applies merged-cell formatting.
   *
   * <p>Consecutive identical values in the same row are merged horizontally.
   *     A column with the same value in all header rows is merged vertically.
   *     Both conditions can apply to a single cell simultaneously.</p>
   *
   * @param sheet the sheet to write the headers into
   */
  public void writeHeaders(Sheet sheet) {
    int poiBasisStartCol = tableStartColumnNumber - 1;
    int poiBasisStartRow = tableStartRowNumber != null ? tableStartRowNumber - 1 : 0;

    // Write cell values.
    for (int rowIdx = 0; rowIdx < headerLabels2d.length; rowIdx++) {
      Row row = sheet.getRow(poiBasisStartRow + rowIdx);
      if (row == null) {
        row = sheet.createRow(poiBasisStartRow + rowIdx);
      }
      for (int colIdx = 0; colIdx < headerLabels2d[rowIdx].length; colIdx++) {
        row.createCell(poiBasisStartCol + colIdx)
            .setCellValue(headerLabels2d[rowIdx][colIdx]);
      }
    }

    // Apply horizontal merges (same label adjacent in same row).
    for (int rowIdx = 0; rowIdx < headerLabels2d.length; rowIdx++) {
      int poiRow = poiBasisStartRow + rowIdx;
      String[] labels = headerLabels2d[rowIdx];
      int startColIdx = 0;
      while (startColIdx < labels.length) {
        int endColIdx = startColIdx;
        while (endColIdx + 1 < labels.length
            && labels[endColIdx + 1].equals(labels[startColIdx])) {
          endColIdx++;
        }
        if (endColIdx > startColIdx) {
          sheet.addMergedRegion(new CellRangeAddress(
              poiRow, poiRow,
              poiBasisStartCol + startColIdx, poiBasisStartCol + endColIdx));
        }
        startColIdx = endColIdx + 1;
      }
    }

    // Apply vertical merges (same label in all rows for a column).
    if (headerLabels2d.length > 1) {
      for (int colIdx = 0; colIdx < headerLabels2d[0].length; colIdx++) {
        String first = headerLabels2d[0][colIdx];
        boolean allSame = true;
        for (int rowIdx = 1; rowIdx < headerLabels2d.length; rowIdx++) {
          if (!headerLabels2d[rowIdx][colIdx].equals(first)) {
            allSame = false;
            break;
          }
        }
        if (allSame) {
          int poiCol = poiBasisStartCol + colIdx;
          sheet.addMergedRegion(new CellRangeAddress(
              poiBasisStartRow, poiBasisStartRow + headerLabels2d.length - 1,
              poiCol, poiCol));
        }
      }
    }
  }
}
