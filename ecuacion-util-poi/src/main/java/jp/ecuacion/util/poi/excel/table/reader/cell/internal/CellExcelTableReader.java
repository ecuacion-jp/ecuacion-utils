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
package jp.ecuacion.util.poi.excel.table.reader.cell.internal;

import jakarta.annotation.Nullable;
import jp.ecuacion.util.poi.excel.table.IfCellExcelTable;
import jp.ecuacion.util.poi.excel.table.reader.core.ExcelTableReader;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;

public abstract class CellExcelTableReader extends ExcelTableReader<Cell>
    implements IfCellExcelTable {

  public CellExcelTableReader(String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);

  }

  @Override
  protected Cell getCellData(Cell cell) {
    return cell;
  }

  @Override
  protected boolean isCellDataEmpty(@Nullable Cell cellData) {
    if (cellData == null) {
      return true;
    }

    String value = new ExcelReadUtil().getStringFromCell(cellData);

    return value == null || value.equals("");
  }
}
