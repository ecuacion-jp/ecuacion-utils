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

import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Provides the excel table writer methods.
 * 
 * <p>Since the number of {@code CellStyle} in an excel file has limit (64,000), 
 * first data line of {@code CellStyle} is reused to cells at the latter lines.</p>
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfExcelTableWriter<T> extends IfExcelTable<T> {

  /**
   * writes cell data to the cell.
   * 
   * @param columnNumberFromZero columnNumberFromZero
   * @param sourceCellData sourceCellData
   * @param destCell destCell
   */
  public void writeToCell(int columnNumberFromZero, T sourceCellData, Cell destCell);
}
