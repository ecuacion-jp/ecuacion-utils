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
package jp.ecuacion.util.poi.excel.table.writer.concrete;

import jakarta.annotation.Nullable;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatFreeExcelTable;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.poi.excel.table.writer.IfDataTypeCellExcelTableWriter;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

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
public class CellFreeExcelTableWriter extends ExcelTableWriter<Cell>
    implements IfFormatFreeExcelTable<Cell>, IfDataTypeCellExcelTableWriter {

  /**
  * Constructs a new instance.
  *
  * @param sheetName See {@link ExcelTable#sheetName}.
  * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
  * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
  */
  public CellFreeExcelTableWriter(@RequireNonnull String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  @Override
  public String getStringValue(@Nullable Cell cellData) {
    return new ExcelReadUtil().getStringFromCell(cellData);
  }

  @Override
  protected void headerCheck(@RequireNonnull Workbook workbook)
      throws EncryptedDocumentException, AppException, IOException {

  }

  private Map<Integer, CellStyle> columnStyleMap = new HashMap<>();
  
  @Override
  public Map<Integer, CellStyle> getColumnStyleMap() {
    return columnStyleMap;
  }
}
