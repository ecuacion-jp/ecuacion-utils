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

import jakarta.annotation.Nullable;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import jp.ecuacion.util.poi.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.poi.excel.table.reader.IfDataTypeStringExcelTableReader;

/**
 * Adds String feature to {@link ExcelTableReader}.
 */
public abstract class StringExcelTableReader extends ExcelTableReader<String>
    implements IfDataTypeStringExcelTableReader {

  protected Map<Integer, DateTimeFormatter> columnDateTimeFormatMap = new HashMap<>();
  
  protected DateTimeFormatter dateTimeFormat;

  /**
   * Constructs a new instance.
   * 
   * {@see ExcelTableReader}
   */
  public StringExcelTableReader(String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }

  /**
   * Sets defaultDateTimeFormat.
   * 
   * @param dateTimeFormat dateTimeFormat string for {@link java.time.format.DateTimeFormatter}.
   * @return StringExcelTableReader (for method chain)
   */
  public StringExcelTableReader defaultDateTimeFormat(DateTimeFormatter dateTimeFormat) {
    this.dateTimeFormat = dateTimeFormat;
    return this;
  }

  /**
   * Sets dateTimeFormat for specific column.
   * 
   * @param columnNumber the column number data is obtained from, 
   *     <b>starting with 1 and column A is equal to columnNumber 1</b>. 
   *     When the far left column of a table is 2 and you want to speciries the far left column,
   *     the columnNumber is 2.
   * @param dateTimeFormat dateTimeFormat string 
   *     for {@link java.time.format.DateTimeFormatter}.
   * @return StringExcelTableReader (for method chain)
   */
  public StringExcelTableReader columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    this.columnDateTimeFormatMap.put(columnNumber, dateTimeFormat);
    return this;
  }

  @Override
  public @Nullable DateTimeFormatter getDateTimeFormat(int columnNumber) {
    return columnDateTimeFormatMap.containsKey(columnNumber)
        ? columnDateTimeFormatMap.get(columnNumber)
        : null;
  }
}
