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
package jp.ecuacion.util.poi.excel.table.writer.cell.internal;

import jp.ecuacion.util.poi.excel.table.writer.core.ExcelTableWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellUtil;

public abstract class CellExcelTableWriter extends ExcelTableWriter<Cell> {

  public CellExcelTableWriter(String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  protected void writeToCell(Cell sourceCellData, Cell destCell) {
    CellCopyPolicy policy = new CellCopyPolicy();
    policy.setCopyCellFormula(false);
    
    CellUtil.copyCell(sourceCellData, destCell, policy, null);
  }
}
