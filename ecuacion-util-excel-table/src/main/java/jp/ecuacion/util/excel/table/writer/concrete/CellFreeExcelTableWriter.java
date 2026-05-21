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
import java.util.HashMap;
import java.util.Map;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatFreeExcelTable;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeCellExcelTableWriter;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with unknown number of columns and no header line.
 *
 * <p>It writes cell values as {@code Cell} object.</p>
 *
 * <p>The header line is not required.
 *     This class writes to the table at the designated position.</p>
 */
public class CellFreeExcelTableWriter extends ExcelTableWriter<Cell>
    implements IfFormatFreeExcelTable<Cell>, IfDataTypeCellExcelTableWriter {

  private boolean copiesDataFormatOnly;

  /**
   * Constructs a new instance with only the sheet name.
   *
   * <p>Defaults: {@code tableStartRowNumber = null}, {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   */
  public CellFreeExcelTableWriter(String sheetName) {
    super(sheetName);
  }

  @Override
  public @Nullable String getStringValue(@Nullable Cell cellData) throws ExcelTableException {
    return ExcelReadUtil.getStringFromCell(cellData, null);
  }

  @Override
  protected void headerCheck(Workbook workbook)
      throws EncryptedDocumentException, IOException {

  }

  private Map<Integer, CellStyle> columnStyleMap = new HashMap<>();

  @Override
  public Map<Integer, CellStyle> getColumnStyleMap() {
    return columnStyleMap;
  }

  @Override
  public boolean copiesDataFormatOnly() {
    return copiesDataFormatOnly;
  }

  @Override
  public CellFreeExcelTableWriter withCopiesDataFormatOnly(boolean copiesDataFormatOnly) {
    this.copiesDataFormatOnly = copiesDataFormatOnly;
    return this;
  }

  @Override
  public CellFreeExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (CellFreeExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public CellFreeExcelTableWriter tableStartColumnNumber(int value) {
    return (CellFreeExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public CellFreeExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellFreeExcelTableWriter) super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellFreeExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (CellFreeExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
