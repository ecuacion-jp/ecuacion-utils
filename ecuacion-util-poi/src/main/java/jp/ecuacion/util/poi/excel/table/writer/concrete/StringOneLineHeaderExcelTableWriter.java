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
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.poi.excel.table.reader.concrete.StringOneLineHeaderExcelTableReader;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.poi.excel.table.writer.IfDataTypeStringExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with known number of columns and one header line.
 *
 * <p>It writes cell values as {@code String}.</p>
 *
 * <p>The header is validated against the expected labels before writing data.</p>
 */
public class StringOneLineHeaderExcelTableWriter extends ExcelTableWriter<String>
    implements IfDataTypeStringExcelTableWriter, IfFormatOneLineHeaderExcelTable<String> {

  private String[] headerLabels;

  /**
   * Constructs a new instance.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels used for header validation.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   *     The row number must specify the header row of the table
   *     since the writer reads and validates the header but does not overwrite it.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public StringOneLineHeaderExcelTableWriter(String sheetName, String[] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.headerLabels = ObjectsUtil.requireNonNull(headerLabels);
  }

  @Override
  public String[] getHeaderLabels() {
    return headerLabels;
  }

  @Override
  protected void headerCheck(Workbook workbook) throws EncryptedDocumentException, IOException {
    new StringOneLineHeaderExcelTableReader(getSheetName(), getHeaderLabelData()[0],
        tableStartRowNumber, tableStartColumnNumber, 1)
            .ignoresAdditionalColumnsOfHeaderData(ignoresAdditionalColumnsOfHeaderData())
            .isVerticalAndHorizontalOpposite(isVerticalAndHorizontalOpposite()).read(workbook);
  }
}
