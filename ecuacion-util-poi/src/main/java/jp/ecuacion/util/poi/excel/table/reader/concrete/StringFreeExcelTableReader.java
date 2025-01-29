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
package jp.ecuacion.util.poi.excel.table.reader.concrete;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.time.format.DateTimeFormatter;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.poi.excel.table.reader.IfFormatFreeExcelTableReader;

/**
 * Reads tables with unknown number of columns, unknown whether it have a header line,
 * unknown header labels if it has a header line.
 * 
 * <p>It obtains cell values as {@code String}.</p>
 *
 * <p>The header line is not necessary. 
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in one line.</p>
 */
public class StringFreeExcelTableReader extends StringExcelTableReader
    implements IfFormatFreeExcelTableReader<String> {

  private NoDataString noDataString;

  /**
   * Constructs a new instance.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link ExcelTableReader#ExcelTableReader(String, Integer, int, Integer, Integer)}.</p>
   */
  public StringFreeExcelTableReader(@RequireNonnull String sheetName,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, @Nullable Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }

  /**
   * Constructs a new instance.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link ExcelTableReader#ExcelTableReader(String, Integer, int, Integer, Integer)}.</p>
   *     
   * @param noDataString noDataString
   */
  public StringFreeExcelTableReader(@RequireNonnull String sheetName,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, @Nullable Integer tableColumnSize,
      @Nonnull NoDataString noDataString) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);

    this.noDataString = noDataString;
  }

  @Override
  public NoDataString getNoDataString() {
    return noDataString;
  }

  @Override
  public StringFreeExcelTableReader defaultDateTimeFormat(DateTimeFormatter dateTimeFormat) {
    return (StringFreeExcelTableReader) super.defaultDateTimeFormat(dateTimeFormat);
  }

  @Override
  public StringFreeExcelTableReader columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringFreeExcelTableReader) super.columnDateTimeFormat(columnNumber, dateTimeFormat);
  }
}
