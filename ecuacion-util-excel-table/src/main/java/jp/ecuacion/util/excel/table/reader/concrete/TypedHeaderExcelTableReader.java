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

import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfDataTypeTypedExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfFormatHeaderExcelTableReader;
import org.jspecify.annotations.Nullable;

/**
 * Reads tables with one or multiple header rows, returning cell values as native Java types.
 *
 * <p>Numeric cells return {@link Double}, date-formatted cells return
 *     {@link java.time.LocalDate} or {@link java.time.LocalDateTime},
 *     string cells return {@link String}, and boolean cells return {@link Boolean}.</p>
 *
 * <p>For the common case of a single header row use
 *     {@link TypedOneLineHeaderExcelTableReader} with a {@code String[]} argument.
 *     For tables with two or more header rows use this class
 *     with a {@code String[][]} argument.</p>
 *
 * <p>Merged cells in the header area are <em>not</em> automatically expanded;
 *     supply the expanded (non-merged) label array explicitly.</p>
 */
public class TypedHeaderExcelTableReader extends ExcelTableReader<Object>
    implements IfFormatHeaderExcelTableReader<Object>, IfDataTypeTypedExcelTableReader {

  private String[][] headerLabels2d;

  /**
   * Constructs a new instance with the sheet name and multiple header rows.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param sheetName See {@link jp.ecuacion.util.excel.table.ExcelTable#sheetName}.
   * @param headerLabels expected labels for each header row: {@code headerLabels[row][col]},
   *     top row first
   */
  public TypedHeaderExcelTableReader(String sheetName, String[][] headerLabels) {
    super(sheetName);
    this.headerLabels2d = ObjectsUtil.requireNonNull(headerLabels);
    setTableColumnSize(getHeaderLabels().length);
  }

  @Override
  public String getFarLeftAndTopHeaderLabel() {
    ObjectsUtil.requireSizeNonZero(headerLabels2d[0]);
    return ObjectsUtil.requireNonNull(headerLabels2d[0][0]);
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
  public List<List<String>> updateAndGetHeaderData(List<List<Object>> excelData)
      throws ExcelTableException {
    List<List<String>> headerRows = new ArrayList<>();
    int numHeaderRows = getNumberOfHeaderLines();
    for (int i = 0; i < numHeaderRows && !excelData.isEmpty(); i++) {
      List<Object> rawRow = excelData.remove(0);
      List<String> strRow = new ArrayList<>();
      for (Object cell : rawRow) {
        strRow.add(getStringValue(cell));
      }
      headerRows.add(strRow);
    }
    return headerRows;
  }

  @Override
  public TypedHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (TypedHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public TypedHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (TypedHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public TypedHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (TypedHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public TypedHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (TypedHeaderExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public TypedHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (TypedHeaderExcelTableReader) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public TypedHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (TypedHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
