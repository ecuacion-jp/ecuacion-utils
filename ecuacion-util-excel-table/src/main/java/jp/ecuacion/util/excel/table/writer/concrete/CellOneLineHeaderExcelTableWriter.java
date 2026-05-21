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

import org.jspecify.annotations.Nullable;

/**
 * Writes tables with a single header row, using {@code Cell} data.
 *
 * <p>This is the recommended class for the common case of a one-line header.
 *     Use this class when you need to preserve cell styles or numeric types.</p>
 */
public class CellOneLineHeaderExcelTableWriter extends CellHeaderExcelTableWriter {

  /**
   * Constructs a new instance with the sheet name and header labels.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName sheet name
   * @param headerLabels expected header labels
   */
  public CellOneLineHeaderExcelTableWriter(String sheetName, String[] headerLabels) {
    super(sheetName, new String[][] {headerLabels});
  }

  @Override
  public CellOneLineHeaderExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableWriter tableStartColumnNumber(int value) {
    return (CellOneLineHeaderExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellOneLineHeaderExcelTableWriter) super.withIgnoresAdditionalColumnsOfHeaderData(
        value);
  }

  @Override
  public CellOneLineHeaderExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (CellOneLineHeaderExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }

  @Override
  public CellOneLineHeaderExcelTableWriter withCopiesDataFormatOnly(boolean copiesDataFormatOnly) {
    return (CellOneLineHeaderExcelTableWriter) super.withCopiesDataFormatOnly(copiesDataFormatOnly);
  }
}
