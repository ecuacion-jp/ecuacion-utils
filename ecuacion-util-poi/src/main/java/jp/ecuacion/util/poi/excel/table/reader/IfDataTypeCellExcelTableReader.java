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
import jp.ecuacion.util.poi.excel.table.IfDataTypeCellExcelTable;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Provides the excel table reader interface 
 *     with object type obtained from the excel data is {@code Cell}.
 */
public interface IfDataTypeCellExcelTableReader
    extends IfDataTypeCellExcelTable, IfExcelTableReader<Cell> {

  @Override
  public default Cell getCellData(Cell cell, int columnNumber) {
    return cell;
  }

  @Override
  public default boolean isCellDataEmpty(@Nullable Cell cellData) {
    if (cellData == null) {
      return true;
    }

    String value = getExcelReadUtil().getStringFromCell(cellData);

    return value == null || value.equals("");
  }
}
