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
 * Reads tables with a single header row, returning cell values as native Java types.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use {@link TypedHeaderExcelTableReader}
 *     with a {@code String[][]} argument.</p>
 */
public class TypedOneLineHeaderExcelTableReader extends TypedHeaderExcelTableReader {

  /**
   * Constructs a new instance with the sheet name and header labels.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param sheetName See {@link jp.ecuacion.util.excel.table.ExcelTable#sheetName}.
   * @param headerLabels expected header labels
   */
  public TypedOneLineHeaderExcelTableReader(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public TypedOneLineHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (TypedOneLineHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    return (TypedOneLineHeaderExcelTableReader) super.withIgnoresAdditionalColumnsOfHeaderData(
        value);
  }

  @Override
  public TypedOneLineHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (TypedOneLineHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
