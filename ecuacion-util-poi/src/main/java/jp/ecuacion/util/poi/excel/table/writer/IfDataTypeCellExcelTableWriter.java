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
package jp.ecuacion.util.poi.excel.table.writer;

import java.util.Map;
import jp.ecuacion.util.poi.excel.table.IfDataTypeCellExcelTable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellUtil;

/**
 * Provides the excel table writer interface 
 *     with object type obtained from the excel data is {@code Cell}.
 */
public interface IfDataTypeCellExcelTableWriter
    extends IfDataTypeCellExcelTable, IfExcelTableWriter<Cell> {

  /**
   * Writes a value to the cell.
   * 
   * @param sourceCellData sourceCellData
   * @param destCell destCell
   */
  public default void writeToCell(int columnNumberFromZero, Cell sourceCellData, Cell destCell) {
    CellCopyPolicy policy = new CellCopyPolicy();
    policy.setCopyCellFormula(false);

    // The number of CellStyle in an excel file has limit: 64,000.
    // If it exceeds, we'll get the exception below. To avoid it CellStyle has to be reused.
    //
    // Exception in thread "main" java.lang.IllegalStateException: The maximum number of Cell Styles
    // was exceeded. You can define up to 64000 style in a .xlsx Workbook
    // 
    // Since CellUtil.copyCell always creates style for each cell 
    // when the source and destination workbook is different, we need to set this to false
    // and override style copy procedure.
    policy.setCopyCellStyle(false);
    
    CellUtil.copyCell(sourceCellData, destCell, policy, null);

    // copy cellStyle
    if (getColumnStyleMap().containsKey(columnNumberFromZero)) {
      destCell.setCellStyle(getColumnStyleMap().get(columnNumberFromZero));
      
    } else {
      destCell.getCellStyle().cloneStyleFrom(sourceCellData.getCellStyle());
      
      getColumnStyleMap().put(columnNumberFromZero, destCell.getCellStyle());
    }
  }
  
  /**
   * Gets {@code columnStyleMap} to reuse {@code CellStyle}.
   * 
   * @return columnStyleMap
   */
  public Map<Integer, CellStyle> getColumnStyleMap();
}
