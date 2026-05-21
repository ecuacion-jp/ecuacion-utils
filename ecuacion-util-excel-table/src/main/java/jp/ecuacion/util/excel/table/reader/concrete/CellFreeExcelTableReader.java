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
package jp.ecuacion.util.excel.table.reader.concrete;

import jp.ecuacion.util.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfDataTypeCellExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfFormatFreeExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;
import org.jspecify.annotations.Nullable;

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
public class CellFreeExcelTableReader extends ExcelTableReader<Cell>
    implements IfFormatFreeExcelTableReader<Cell>, IfDataTypeCellExcelTableReader {

  /**
   * Constructs a new instance with only the sheet name.
   *
   * <p>Defaults: {@code tableStartRowNumber = null}, {@code tableStartColumnNumber = 1},
   *     {@code tableRowSize = null}, {@code tableColumnSize = null}.</p>
   *
   * @param sheetName See {@link jp.ecuacion.util.excel.table.ExcelTable#sheetName}.
   */
  public CellFreeExcelTableReader(String sheetName) {
    super(sheetName);
  }

  @Override
  public CellFreeExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (CellFreeExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public CellFreeExcelTableReader tableStartColumnNumber(int value) {
    return (CellFreeExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public CellFreeExcelTableReader tableRowSize(@Nullable Integer value) {
    return (CellFreeExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public CellFreeExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (CellFreeExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public CellFreeExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellFreeExcelTableReader) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellFreeExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (CellFreeExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
