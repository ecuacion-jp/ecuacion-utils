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
package jp.ecuacion.util.poi.excel.table.writer.cell;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.util.poi.excel.table.IfCellExcelTable;
import jp.ecuacion.util.poi.excel.table.IfOneLineHeaderFormatExcelTable;
import jp.ecuacion.util.poi.excel.table.reader.string.StringFreeFormatExcelTableReader;
import jp.ecuacion.util.poi.excel.table.writer.cell.internal.CellExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;

/**
 * writer.
 */
public class CellOneLineHeaderFormatExcelTableWriter extends CellExcelTableWriter
    implements IfCellExcelTable, IfOneLineHeaderFormatExcelTable<Cell> {

  @Nonnull
  private String[] headerLabels;

  /**
   * Constructs a new instance.
   * 
   * @param sheetName sheetName
   * @param headerLabels headerLabels
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   */
  public CellOneLineHeaderFormatExcelTableWriter(@RequireNonnull String sheetName,
      @Nonnull String[] headerLabels, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.headerLabels = headerLabels;
  }

  @Override
  @Nonnull
  public String[] getHeaderLabels() {
    return headerLabels;
  }

  @Override
  protected List<List<String>> getHeaderList(String templateFilePath, int tableColumnSize)
      throws EncryptedDocumentException, AppException, IOException {

    return new StringFreeFormatExcelTableReader(getSheetName(), tableStartRowNumber,
        tableStartColumnNumber, 1, tableColumnSize).read(templateFilePath.toString());
  }
}
