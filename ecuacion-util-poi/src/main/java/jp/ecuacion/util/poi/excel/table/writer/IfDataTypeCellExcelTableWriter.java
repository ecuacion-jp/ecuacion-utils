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
   * <p>The number of CellStyle in an excel file has limit: 64,000.
   *     If it exceeds, we'll get the exception below. To avoid it CellStyle has to be reused.
   *     To reuse style</p>
   * 
   * @param columnNumberFromZero column number starting from zero. 
   *     Used as a key to store and reuse the style for each column.
   * @param sourceCellData sourceCellData
   * @param destCell destCell
   */
  public default void writeToCell(int columnNumberFromZero, Cell sourceCellData, Cell destCell) {
    CellCopyPolicy policy = new CellCopyPolicy();
    policy.setCopyCellFormula(false);

    // Since CellUtil.copyCell always creates style for each cell
    // when the source and destination workbook is different, we need to set this to false
    // and override style copy procedure.
    policy.setCopyCellStyle(false);

    CellUtil.copyCell(sourceCellData, destCell, policy, null);

    // copy cellStyle
    if (getColumnStyleMap().containsKey(columnNumberFromZero)) {
      destCell.setCellStyle(getColumnStyleMap().get(columnNumberFromZero));

    } else {
      if (sourceCellData != null) {
        if (copiesDataFormatOnly()) {
          destCell.getCellStyle().setDataFormat(sourceCellData.getCellStyle().getDataFormat());

        } else {
          // Under some conditions org.apache.xmlbeans.impl.vales.XmlValueDisconnectedException
          // occurs
          // when workbook.save() is called after cloneStyleFrom is used.
          // Reason is unclear but it seems to happen when the java object (like Cells) exists
          // but the xml in workbook is gone.
          // So I think it's may be because of something related to the cloned style
          // which does not exist in xml.
          // That's why I put createCellStyle() before cloneStyleFrom()
          // and problem seems to be resolved.
          destCell.setCellStyle(destCell.getRow().getSheet().getWorkbook().createCellStyle());
          destCell.getCellStyle().cloneStyleFrom(sourceCellData.getCellStyle());
        }

        getColumnStyleMap().put(columnNumberFromZero, destCell.getCellStyle());
      }
    }
  }
  
  /**
   * Sets copiesDataFormatOnly value.
   * 
   * <p>When this is {@code true}, whole style is not copied 
   *     to the destination cell, but {@code DataFormat} only. <br>
   *     This means grid-line, font, font-size, cell color, etc... of the cell is not copied.</p>
   *     
   * @param value boolean
   *     
   */
  public IfDataTypeCellExcelTableWriter copiesDataFormatOnly(boolean value);
  
  /**
   * Returns copiesDataFormatOnly value.
   * 
   * <p>When this is {@code true}, whole style is not copied 
   *     to the destination cell, but {@code DataFormat} only. <br>
   *     This means grid-line, font, font-size, cell color, etc... of the cell is not copied.</p>
   *     
   * @return boolean
   */
  public boolean copiesDataFormatOnly();

  /**
   * Gets {@code columnStyleMap} to reuse {@code CellStyle}.
   * 
   * @return columnStyleMap
   */
  public Map<Integer, CellStyle> getColumnStyleMap();
}
