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
package jp.ecuacion.util.poi.read.cell.reader.internal;

import jakarta.annotation.Nullable;
import jp.ecuacion.util.poi.read.core.reader.TableReader;
import jp.ecuacion.util.poi.util.PoiReadUtil;
import org.apache.poi.ss.usermodel.Cell;

public abstract class CellTableReader extends TableReader<Cell> {

  public CellTableReader(String sheetName, Integer tableStartRowNumber, int tableStartColumnNumber,
      Integer tableRowSize, Integer tableColumnSize) {
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
    
    String value = new PoiReadUtil().getStringFromCell(cellData);

    return value == null || value.equals("");
  }
}
