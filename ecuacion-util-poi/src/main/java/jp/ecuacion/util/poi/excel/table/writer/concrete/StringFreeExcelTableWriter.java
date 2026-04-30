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

import java.io.IOException;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatFreeExcelTable;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.poi.excel.table.writer.IfDataTypeStringExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with unknown number of columns and no header line.
 *
 * <p>It writes cell values as {@code String}.</p>
 *
 * <p>The header line is not required.
 *     This class writes to the table at the designated position.</p>
 */
public class StringFreeExcelTableWriter extends ExcelTableWriter<String>
    implements IfFormatFreeExcelTable<String>, IfDataTypeStringExcelTableWriter {

  /**
   * Constructs a new instance.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public StringFreeExcelTableWriter(String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  @Override
  protected void headerCheck(Workbook workbook) throws EncryptedDocumentException, IOException {
    // No header to check for free-format tables.
  }
}
