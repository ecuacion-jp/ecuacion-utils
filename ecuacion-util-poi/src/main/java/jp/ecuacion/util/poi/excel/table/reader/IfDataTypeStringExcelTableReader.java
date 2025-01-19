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
package jp.ecuacion.util.poi.excel.table.reader;

import jakarta.annotation.Nullable;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.IfDataTypeStringExcelTable;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Provides the excel table reader interface 
 *     with object type obtained from the excel data is {@code String}.
 */
public interface IfDataTypeStringExcelTableReader
    extends IfDataTypeStringExcelTable, IfExcelTableReader<String> {

  @Override
  public default String getCellData(Cell cell, int columnNumber) {
    return getExcelReadUtil().getStringFromCell(cell, getDateTimeFormat(columnNumber));
  }

  @Override
  public default boolean isCellDataEmpty(@Nullable String cellData) {
    return StringUtils.isEmpty(cellData);
  }

  /**
   * Gets NoDataString.
   * 
   * @return NoDataString
   */
  public NoDataString getNoDataString();

  /**
   * Returns the date time format considering {@code columnDateTimeFormatMap}
   *     and {@code defaultDateTimeFormat}.
   * 
   * @param columnNumber the column number data is obtained from, 
   *     <b>starting with 1 and column A is equal to columnNumber 1</b>. 
   *     When the far left column of a table is 2 and you want to speciries the far left column,
   *     the columnNumber is 2.
   * @return date format
   */
  public String getDateTimeFormat(int columnNumber);
}
