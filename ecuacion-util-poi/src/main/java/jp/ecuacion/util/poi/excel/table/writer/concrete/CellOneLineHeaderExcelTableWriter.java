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

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.io.IOException;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.poi.excel.table.reader.concrete.StringOneLineHeaderExcelTableReader;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.poi.excel.table.writer.IfDataTypeCellExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Reads tables with known number of columns, one line header labels if it has a header line.
 */
public class CellOneLineHeaderExcelTableWriter extends ExcelTableWriter<Cell>
    implements IfDataTypeCellExcelTableWriter, IfFormatOneLineHeaderExcelTable<Cell> {

  @Nonnull
  private String[] headerLabels;

  /**
   * Constructs a new instance.
   * 
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public CellOneLineHeaderExcelTableWriter(@RequireNonnull String sheetName,
      @RequireNonnull String[] headerLabels, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.headerLabels = ObjectsUtil.paramRequireNonNull(headerLabels);
  }

  @Override
  @Nonnull
  public String[] getHeaderLabels() {
    return headerLabels;
  }

  @Override
  protected void headerCheck(@RequireNonnull Workbook workbook)
      throws EncryptedDocumentException, AppException, IOException {

    new StringOneLineHeaderExcelTableReader(getSheetName(), getHeaderLabelData()[0],
        tableStartRowNumber, tableStartColumnNumber, 1).read(workbook);
  }
}
