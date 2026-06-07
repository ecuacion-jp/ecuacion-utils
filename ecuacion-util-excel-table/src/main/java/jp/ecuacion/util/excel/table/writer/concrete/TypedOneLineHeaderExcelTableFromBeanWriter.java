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

import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;

/**
 * Writes tables with a single header row from a list of {@link TypedExcelTableBean} instances.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     For tables with two or more header rows use
 *     {@link TypedHeaderExcelTableFromBeanWriter} with a {@code String[][]} argument.</p>
 *
 * @param <T> the bean type, must extend {@link TypedExcelTableBean}
 */
public class TypedOneLineHeaderExcelTableFromBeanWriter<T extends TypedExcelTableBean>
    extends TypedHeaderExcelTableFromBeanWriter<T> {

  /**
   * Constructs a new instance with the sheet name and a single header row.
   *
   * @param sheetName sheet name
   * @param headerLabels expected header labels for the single header row
   */
  public TypedOneLineHeaderExcelTableFromBeanWriter(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public TypedOneLineHeaderExcelTableFromBeanWriter<T> defaultDateFormat(String formatPattern) {
    super.defaultDateFormat(formatPattern);
    return this;
  }

  @Override
  public TypedOneLineHeaderExcelTableFromBeanWriter<T> defaultDateTimeFormat(
      String formatPattern) {
    super.defaultDateTimeFormat(formatPattern);
    return this;
  }
}
