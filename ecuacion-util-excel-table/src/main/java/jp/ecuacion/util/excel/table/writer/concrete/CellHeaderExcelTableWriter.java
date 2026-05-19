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
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfFormatHeaderExcelTable;
import jp.ecuacion.util.excel.table.reader.concrete.StringHeaderExcelTableReader;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter;
import jp.ecuacion.util.excel.table.writer.IfDataTypeCellExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with known number of columns and one or more header rows.
 *
 * <p>The header is validated against the expected labels before writing data.</p>
 */
public class CellHeaderExcelTableWriter extends ExcelTableWriter<Cell>
    implements IfDataTypeCellExcelTableWriter, IfFormatHeaderExcelTable<Cell> {

  private boolean copiesDataFormatOnly;

  private String[] headerLabels;

  /**
   * Constructs a new instance with the sheet name and header labels.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected header labels
   */
  public CellHeaderExcelTableWriter(String sheetName, String[] headerLabels) {
    super(sheetName);
    this.headerLabels = ObjectsUtil.requireNonNull(headerLabels);
  }

  /**
   * Constructs a new instance.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   *     The row number must specify the header row of the table
   *     Since the writer does not overwrite the header, but the writer does read and validate it.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   *
   * @deprecated Use the minimal constructor with fluent setters instead.
   */
  @Deprecated
  public CellHeaderExcelTableWriter(String sheetName,
      String[] headerLabels, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.headerLabels = ObjectsUtil.requireNonNull(headerLabels);
  }

  @Override
  public String[] getHeaderLabels() {
    return headerLabels;
  }

  @Override
  protected void headerCheck(Workbook workbook)
      throws EncryptedDocumentException, IOException {

    new StringHeaderExcelTableReader(getSheetName(), getHeaderLabelData()[0])
        .tableStartRowNumber(tableStartRowNumber)
        .tableStartColumnNumber(tableStartColumnNumber)
        .tableRowSize(1)
        .withIgnoresAdditionalColumnsOfHeaderData(ignoresAdditionalColumnsOfHeaderData())
        .withVerticalAndHorizontalOpposite(isVerticalAndHorizontalOpposite()).read(workbook);
  }

  private Map<Integer, CellStyle> columnStyleMap = new HashMap<>();

  @Override
  public Map<Integer, CellStyle> getColumnStyleMap() {
    return columnStyleMap;
  }

  @SuppressWarnings("InlineMeSuggester")
  @Override
  @Deprecated
  public CellHeaderExcelTableWriter copiesDataFormatOnly(boolean copiesDataFormatOnly) {
    return withCopiesDataFormatOnly(copiesDataFormatOnly);
  }

  @Override
  public boolean copiesDataFormatOnly() {
    return copiesDataFormatOnly;
  }

  @Override
  public CellHeaderExcelTableWriter withCopiesDataFormatOnly(boolean copiesDataFormatOnly) {
    this.copiesDataFormatOnly = copiesDataFormatOnly;
    return this;
  }

  @Override
  public CellHeaderExcelTableWriter tableStartRowNumber(@Nullable Integer value) {
    return (CellHeaderExcelTableWriter) super.tableStartRowNumber(value);
  }

  @Override
  public CellHeaderExcelTableWriter tableStartColumnNumber(int value) {
    return (CellHeaderExcelTableWriter) super.tableStartColumnNumber(value);
  }

  @Override
  public CellHeaderExcelTableWriter withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellHeaderExcelTableWriter)
        super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellHeaderExcelTableWriter withVerticalAndHorizontalOpposite(boolean value) {
    return (CellHeaderExcelTableWriter) super.withVerticalAndHorizontalOpposite(value);
  }
}
