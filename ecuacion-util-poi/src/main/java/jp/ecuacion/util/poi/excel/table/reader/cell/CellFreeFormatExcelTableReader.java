/*
 * Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package jp.ecuacion.util.poi.excel.table.reader.cell;

import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.excel.table.reader.cell.internal.CellExcelTableReader;
import jp.ecuacion.util.poi.excel.table.reader.core.IfFreeFormatExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Reads tables with unknown number of columns, unknown whether it have a header line,
 * unknown header labels if it has a header line.
 * 
 * <p>It obtains cell values as {@code Cell} object.</p>
 * 
 * <p>The header line is not necessary. 
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in one line.</p>
 */
public class CellFreeFormatExcelTableReader extends CellExcelTableReader
    implements IfFreeFormatExcelTableReader<Cell> {

  /**
   * Constructs a new instance.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link jp.ecuacion.util.poi.excel.table.reader.core.ExcelTableReader#ExcelTableReader(
   *     String, Integer, int, Integer, Integer)}.</p>
   */
  public CellFreeFormatExcelTableReader(@RequireNonnull String sheetName,
      Integer tableStartRowNumber, int tableStartColumnNumber, Integer tableRowSize,
      Integer tableColumnSize) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }
}
