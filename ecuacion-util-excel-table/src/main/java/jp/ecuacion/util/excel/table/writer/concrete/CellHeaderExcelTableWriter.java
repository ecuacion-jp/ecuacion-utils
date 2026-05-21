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
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatHeaderExcelTable;
import jp.ecuacion.util.excel.table.reader.concrete.StringHeaderExcelTableReader;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeCellExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with one or multiple header rows, using {@code Cell} data.
 *
 * <p>For the common case of a single header row use
 *     {@link CellOneLineHeaderExcelTableWriter} with a {@code String[]} argument.
 *     For tables with two or more header rows use this class
 *     with a {@code String[][]} argument.</p>
 *
 * <p>The header in the template file is validated against {@code headerLabels} before writing.</p>
 */
public class CellHeaderExcelTableWriter extends ExcelTableWriter<Cell>
    implements IfDataTypeCellExcelTableWriter, IfFormatHeaderExcelTable<Cell> {

  private boolean copiesDataFormatOnly;

  /** All header rows' labels: {@code headerLabels2d[row][col]}. */
  private String[][] headerLabels2d;

  /**
   * Constructs a new instance with the sheet name and multiple header rows.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected labels for each header row: {@code headerLabels[row][col]},
   *     top row first. All rows must have the same length.
   */
  public CellHeaderExcelTableWriter(String sheetName, String[][] headerLabels) {
    super(sheetName);
    this.headerLabels2d = ObjectsUtil.requireNonNull(headerLabels);
  }

  @Override
  public String[] getHeaderLabels() {
    return headerLabels2d[headerLabels2d.length - 1];
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
  protected void headerCheck(Workbook workbook)
      throws EncryptedDocumentException, IOException {
    new StringHeaderExcelTableReader(getSheetName(), headerLabels2d)
        .tableStartRowNumber(tableStartRowNumber)
        .tableStartColumnNumber(tableStartColumnNumber)
        .tableRowSize(1)
        .withIgnoresAdditionalColumnsOfHeaderData(ignoresAdditionalColumnsOfHeaderData())
        .withVerticalAndHorizontalOpposite(isVerticalAndHorizontalOpposite()).read(workbook);
  }

  private Map<Integer, CellStyle> columnStyleMap = new HashMap<>();

  @Override
  public Map<Integer, CellStyle> getColumnStyleMap() {
    return columnStyleMap;
  }

  @Override
  public boolean copiesDataFormatOnly() {
    return copiesDataFormatOnly;
  }

  @Override
  public CellHeaderExcelTableWriter withCopiesDataFormatOnly(boolean copiesDataFormatOnly) {
    this.copiesDataFormatOnly = copiesDataFormatOnly;
    return this;
  }

  @Override
  public CellHeaderExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (CellHeaderExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public CellHeaderExcelTableWriter tableStartColumnNumber(int value) {
    return (CellHeaderExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public CellHeaderExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellHeaderExcelTableWriter) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellHeaderExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (CellHeaderExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
