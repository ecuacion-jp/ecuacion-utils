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
package jp.ecuacion.util.excel.table.writer.concrete;

import java.io.IOException;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatFreeExcelTable;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeStringExcelTableWriter;
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
   * Constructs a new instance with only the sheet name.
   *
   * <p>Defaults: {@code tableStartRowNumber = null}, {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   */
  public StringFreeExcelTableWriter(String sheetName) {
    super(sheetName);
  }

  @Override
  protected void headerCheck(Workbook workbook) throws EncryptedDocumentException, IOException {
    // No header to check for free-format tables.
  }

  @Override
  public StringFreeExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (StringFreeExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public StringFreeExcelTableWriter tableStartColumnNumber(int value) {
    return (StringFreeExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public StringFreeExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (StringFreeExcelTableWriter) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public StringFreeExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (StringFreeExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
