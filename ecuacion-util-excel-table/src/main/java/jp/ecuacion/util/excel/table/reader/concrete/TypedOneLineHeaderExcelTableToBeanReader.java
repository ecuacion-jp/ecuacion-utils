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

import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;
import org.jspecify.annotations.Nullable;

/**
 * Reads a single-header-row Excel table and stores each data row into a bean.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use
 *     {@link TypedHeaderExcelTableToBeanReader} with a {@code String[][]} argument.</p>
 *
 * @param <T> the bean type, must extend {@link TypedExcelTableBean}
 */
public class TypedOneLineHeaderExcelTableToBeanReader<T extends TypedExcelTableBean>
    extends TypedHeaderExcelTableToBeanReader<T> {

  /**
   * Constructs a new instance with a single header row.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param beanClass the class of the bean ({@code T}) — pass explicitly because Java generics
   *     do not allow {@code T.class}
   * @param sheetName sheet name
   * @param headerLabels expected header labels for the single header row
   */
  public TypedOneLineHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[] headerLabels) {
    super(beanClass, sheetName, new String[][] {headerLabels});
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> tableStartRowNumber(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super.tableStartRowNumber(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> tableStartColumnNumber(int value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super.tableStartColumnNumber(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> tableRowSize(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super.tableRowSize(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> tableColumnSize(@Nullable Integer value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super.tableColumnSize(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super
        .withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public TypedOneLineHeaderExcelTableToBeanReader<T> withVerticalAndHorizontalOpposite(
      boolean value) {
    return (TypedOneLineHeaderExcelTableToBeanReader<T>) super.withVerticalAndHorizontalOpposite(
        value);
  }
}
