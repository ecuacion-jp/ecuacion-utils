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

import org.jspecify.annotations.Nullable;

/**
 * Reads tables with a single header row, returning cell values as {@code Cell}.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     Use this class when you need access to cell styles or numeric types in addition
 *     to the cell value.</p>
 */
public class CellOneLineHeaderExcelTableReader extends CellHeaderExcelTableReader {

  /**
   * Constructs a new instance with the sheet name and header labels.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param sheetName See {@link jp.ecuacion.util.excel.table.ExcelTable#sheetName}.
   * @param headerLabels expected header labels
   */
  public CellOneLineHeaderExcelTableReader(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (CellOneLineHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    return (CellOneLineHeaderExcelTableReader) super.withIgnoresAdditionalColumnsOfHeaderData(
        value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (CellOneLineHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
