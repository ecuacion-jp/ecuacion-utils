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

import jp.ecuacion.util.excel.table.ExcelTable;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with a single header row, using {@code String} data.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use
 *     {@link StringHeaderExcelTableWriter} with a {@code String[][]} argument.</p>
 *
 * <p>Internally delegates to {@link StringHeaderExcelTableWriter} with the
 *     provided labels wrapped in a single-element {@code String[][]}.</p>
 */
public class StringOneLineHeaderExcelTableWriter extends StringHeaderExcelTableWriter {

  /**
   * Constructs a new instance with the sheet name and a single header row.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels for the single header row
   */
  public StringOneLineHeaderExcelTableWriter(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public StringOneLineHeaderExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableWriter tableStartColumnNumber(int value) {
    return (StringOneLineHeaderExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    return (StringOneLineHeaderExcelTableWriter) super
        .withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public StringOneLineHeaderExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (StringOneLineHeaderExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
