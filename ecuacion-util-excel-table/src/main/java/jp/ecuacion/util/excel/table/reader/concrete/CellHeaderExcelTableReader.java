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
import jp.ecuacion.util.excel.table.reader.IfDataTypeCellExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfFormatHeaderExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;
import org.jspecify.annotations.Nullable;

/**
 * Reads tables with one or multiple header rows, returning cell values as {@code Cell}.
 *
 * <p>For the common case of a single header row use
 *     {@link CellOneLineHeaderExcelTableReader} with a {@code String[]} argument.
 *     For tables with two or more header rows use this class
 *     with a {@code String[][]} argument.</p>
 *
 * <p>The last header row determines the column count.
 *     Merged cells in the header area are <em>not</em> automatically expanded;
 *     supply the expanded (non-merged) label array explicitly.</p>
 */
public class CellHeaderExcelTableReader extends ExcelTableReader<Cell>
    implements IfFormatHeaderExcelTableReader<Cell>, IfDataTypeCellExcelTableReader {

  /** All header rows' labels: {@code headerLabels2d[row][col]}. */
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
  public CellHeaderExcelTableReader(String sheetName, String[][] headerLabels) {
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
  public List<List<String>> updateAndGetHeaderData(List<List<Cell>> excelData)
      throws ExcelTableException {
    List<List<String>> headerRows = new ArrayList<>();
    int numHeaderRows = getNumberOfHeaderLines();
    for (int i = 0; i < numHeaderRows && !excelData.isEmpty(); i++) {
      List<Cell> rawRow = excelData.remove(0);
      List<String> strRow = new ArrayList<>();
      for (Cell cell : rawRow) {
        strRow.add(getStringValue(cell));
      }
      headerRows.add(strRow);
    }
    return headerRows;
  }

  @Override
  public CellHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (CellHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public CellHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (CellHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public CellHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (CellHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public CellHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (CellHeaderExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public CellHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellHeaderExcelTableReader) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (CellHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
