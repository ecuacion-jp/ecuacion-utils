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
import jp.ecuacion.util.excel.table.bean.StringExcelTableBean;
import org.jspecify.annotations.Nullable;

/**
 * Reads a single-header-row Excel table and stores each data row into a bean.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use
 *     {@link StringHeaderExcelTableToBeanReader} with a {@code String[][]} argument.</p>
 *
 * <p>Internally delegates to {@link StringHeaderExcelTableToBeanReader} with the
 *     provided labels wrapped in a single-element {@code String[][]}.</p>
 *
 * @param <T> the bean type, must extend {@link StringExcelTableBean}
 */
public class StringOneLineHeaderExcelTableToBeanReader<T extends StringExcelTableBean>
    extends StringHeaderExcelTableToBeanReader<T> {

  /**
   * Constructs a new instance with a single header row.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param beanClass the class of the bean ({@code T}) — pass explicitly because Java generics
   *     do not allow {@code T.class}
   * @param sheetName sheet name
   * @param headerLabels expected header labels
   */
  public StringOneLineHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[] headerLabels) {
    super(beanClass, sheetName, new String[][] {headerLabels});
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> defaultDateTimeFormat(
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>)
        super.defaultDateTimeFormat(dateTimeFormat);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.columnDateTimeFormat(columnNumber,
        dateTimeFormat);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> withVerticalAndHorizontalOpposite(
      boolean value) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.withVerticalAndHorizontalOpposite(
        value);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> noDataString(NoDataString noDataString) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.noDataString(noDataString);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> tableStartRowNumber(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.tableStartRowNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> tableStartColumnNumber(int value) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.tableStartColumnNumber(value);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> tableRowSize(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.tableRowSize(value);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> tableColumnSize(@Nullable Integer value) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.tableColumnSize(value);
  }
}
