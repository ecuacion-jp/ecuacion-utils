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
package jp.ecuacion.util.poi.excel.table.reader.string;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.reader.core.IfOneLineHeaderFormatExcelTableReader;
import jp.ecuacion.util.poi.excel.table.reader.string.internal.StringExcelTableReader;

/**
 * Reads tables with known number of columns, known header labels 
 * and known start position of the table.
 * 
 * <p>It obtains cell values as {@code String}.</p>
 * 
 * <p>The header line is required.
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in a line.</p>
 */
public class StringOneLineHeaderFormatExcelTableReader extends StringExcelTableReader
    implements IfOneLineHeaderFormatExcelTableReader<String> {

  @Nonnull
  private String[] headerLabels;

  /**
   * Constructs a new instance. the obtained value 
   *     from an empty cell is {@code null}.
   *     
   * <p>{@code tableColumnSize} is not designated 
   *     because {@code tableColumnSize} of the table is obviously equal to
   *     the length of the header array.</p>
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link jp.ecuacion.util.poi.excel.table.reader.core.ExcelTableReader#ExcelTableReader(
   *     String, Integer, int, Integer, Integer)}.</p>
   */
  public StringOneLineHeaderFormatExcelTableReader(@RequireNonnull String sheetName,
      @Nonnull String[] headerLabels, Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize) {
    this(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        NoDataString.NULL);
  }

  /**
   * Constructs a new instance with the obtained value from an empty cell.
   * 
   * <p>{@code tableColumnSize} is not designated 
   *     because {@code tableColumnSize} of the table is obviously equal to
   *     the length of the header array.</p>
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link jp.ecuacion.util.poi.excel.table.reader.core.ExcelTableReader#ExcelTableReader(
   *     String, Integer, int, Integer, Integer)}.</p>
   *     
   * @param sheetName the sheet name of the excel file
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableRowSize tableRowSize
   * @param noDataString the obtained value from an empty cell. {@code null} or {@code ""}.
   */
  public StringOneLineHeaderFormatExcelTableReader(@RequireNonnull String sheetName,
      @Nonnull String[] headerLabels, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber, @Nullable Integer tableRowSize,
      @Nonnull NoDataString noDataString) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, null, noDataString);

    this.headerLabels = headerLabels;

    // Since "Cannot refer to an instance method while explicitly invoking a constructor",
    // First set "null" in "super(...)" and then set the actual value here.
    setTableColumnSize(getHeaderLabels().length);
  }

  @Override
  @Nonnull
  public String[] getHeaderLabels() {
    return headerLabels;
  }
}
