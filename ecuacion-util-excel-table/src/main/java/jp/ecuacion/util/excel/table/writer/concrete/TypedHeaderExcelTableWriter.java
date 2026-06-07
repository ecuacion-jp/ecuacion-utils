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
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatHeaderExcelTable;
import jp.ecuacion.util.excel.table.reader.concrete.TypedHeaderExcelTableReader;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeTypedExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with one or multiple header rows, using native Java type values.
 *
 * <p>{@link java.time.LocalDate} and {@link java.time.LocalDateTime} values are written as
 *     date-formatted cells, {@link Number} values as numeric cells, {@link Boolean} values as
 *     boolean cells, and any other value via {@link Object#toString()}.
 *     See {@link IfDataTypeTypedExcelTableWriter} for details.</p>
 *
 * <p>The header in the template file is validated against {@code headerLabels} before writing.</p>
 */
public class TypedHeaderExcelTableWriter extends ExcelTableWriter<Object>
    implements IfDataTypeTypedExcelTableWriter, IfFormatHeaderExcelTable<Object> {

  private String[][] headerLabels2d;

  private String dateFormat = "yyyy-mm-dd";
  private String dateTimeFormat = "yyyy-mm-dd hh:mm:ss";
  private Map<String, CellStyle> dateCellStyleMap = new HashMap<>();

  /**
   * Constructs a new instance with the sheet name and multiple header rows.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels: {@code headerLabels[row][col]}, top row first.
   *     All rows must have the same length.
   */
  public TypedHeaderExcelTableWriter(String sheetName, String[][] headerLabels) {
    super(sheetName);
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
   * Validates the template file's headers using a {@link TypedHeaderExcelTableReader}.
   *
   * @param workbook workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  @Override
  protected void headerCheck(Workbook workbook) throws EncryptedDocumentException, IOException {
    new TypedHeaderExcelTableReader(getSheetName(), headerLabels2d)
        .tableStartRowNumber(tableStartRowNumber).tableStartColumnNumber(tableStartColumnNumber)
        .tableRowSize(1)
        .withIgnoresAdditionalColumnsOfHeaderData(ignoresAdditionalColumnsOfHeaderData())
        .withVerticalAndHorizontalOpposite(isVerticalAndHorizontalOpposite()).read(workbook);
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
    int poiBasisStartRow =
        tableStartRowNumber != null ? Objects.requireNonNull(tableStartRowNumber) - 1 : 0;

    // Write cell values.
    for (int rowIdx = 0; rowIdx < headerLabels2d.length; rowIdx++) {
      Row row = sheet.getRow(poiBasisStartRow + rowIdx);
      if (row == null) {
        row = sheet.createRow(poiBasisStartRow + rowIdx);
      }
      for (int colIdx = 0; colIdx < headerLabels2d[rowIdx].length; colIdx++) {
        row.createCell(poiBasisStartCol + colIdx).setCellValue(headerLabels2d[rowIdx][colIdx]);
      }
    }

    // Apply horizontal merges (same label adjacent in same row).
    for (int rowIdx = 0; rowIdx < headerLabels2d.length; rowIdx++) {
      int poiRow = poiBasisStartRow + rowIdx;
      String[] labels = headerLabels2d[rowIdx];
      int startColIdx = 0;
      while (startColIdx < labels.length) {
        int endColIdx = startColIdx;
        while (endColIdx + 1 < labels.length && labels[endColIdx + 1].equals(labels[startColIdx])) {
          endColIdx++;
        }
        if (endColIdx > startColIdx) {
          sheet.addMergedRegion(new CellRangeAddress(poiRow, poiRow, poiBasisStartCol + startColIdx,
              poiBasisStartCol + endColIdx));
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
          sheet.addMergedRegion(new CellRangeAddress(poiBasisStartRow,
              poiBasisStartRow + headerLabels2d.length - 1, poiCol, poiCol));
        }
      }
    }
  }

  @Override
  public String getDateFormat() {
    return dateFormat;
  }

  @Override
  public String getDateTimeFormat() {
    return dateTimeFormat;
  }

  @Override
  public Map<String, CellStyle> getDateCellStyleMap() {
    return dateCellStyleMap;
  }

  /**
   * Sets the Excel number-format pattern used for {@link java.time.LocalDate} values when the
   *     destination cell does not already have a date format. Defaults to {@code "yyyy-mm-dd"}.
   *
   * @param formatPattern the Excel number-format pattern
   * @return this writer
   */
  public TypedHeaderExcelTableWriter defaultDateFormat(String formatPattern) {
    this.dateFormat = formatPattern;
    return this;
  }

  /**
   * Sets the Excel number-format pattern used for {@link java.time.LocalDateTime} values when the
   *     destination cell does not already have a date format.
   *     Defaults to {@code "yyyy-mm-dd hh:mm:ss"}.
   *
   * @param formatPattern the Excel number-format pattern
   * @return this writer
   */
  public TypedHeaderExcelTableWriter defaultDateTimeFormat(String formatPattern) {
    this.dateTimeFormat = formatPattern;
    return this;
  }

  @Override
  public TypedHeaderExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (TypedHeaderExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public TypedHeaderExcelTableWriter tableStartColumnNumber(int value) {
    return (TypedHeaderExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public TypedHeaderExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (TypedHeaderExcelTableWriter) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public TypedHeaderExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (TypedHeaderExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
