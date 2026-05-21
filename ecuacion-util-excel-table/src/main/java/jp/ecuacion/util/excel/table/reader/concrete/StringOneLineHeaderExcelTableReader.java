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

import java.time.format.DateTimeFormatter;
import jp.ecuacion.util.excel.enums.NoDataString;
import org.jspecify.annotations.Nullable;

/**
 * Reads tables with a single header row, returning cell values as {@code String}.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use
 *     {@link StringHeaderExcelTableReader} with a {@code String[][]} argument.</p>
 *
 * <p>Internally delegates to {@link StringHeaderExcelTableReader} with the
 *     provided labels wrapped in a single-element {@code String[][]}.</p>
 */
public class StringOneLineHeaderExcelTableReader extends StringHeaderExcelTableReader {

  /**
   * Constructs a new instance with the sheet name and a single header row.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null},
   *     {@code noDataString = NoDataString.NULL}.</p>
   *
   * @param sheetName sheet name
   * @param headerLabels expected header labels for the single header row
   */
  public StringOneLineHeaderExcelTableReader(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public StringOneLineHeaderExcelTableReader noDataString(NoDataString noDataString) {
    return (StringOneLineHeaderExcelTableReader) super.noDataString(noDataString);
  }

  @Override
  public StringOneLineHeaderExcelTableReader defaultDateTimeFormat(
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableReader) super.defaultDateTimeFormat(dateTimeFormat);
  }

  @Override
  public StringOneLineHeaderExcelTableReader columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableReader) super.columnDateTimeFormat(columnNumber,
        dateTimeFormat);
  }

  @Override
  public StringOneLineHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    return (StringOneLineHeaderExcelTableReader) super
        .withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public StringOneLineHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (StringOneLineHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }

  @Override
  public StringOneLineHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (StringOneLineHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public StringOneLineHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableReader) super.tableColumnSize(value);
  }
}
