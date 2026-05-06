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

import java.util.Objects;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfDataTypeCellExcelTableReader;
import jp.ecuacion.util.excel.table.reader.IfFormatOneLineHeaderExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;
import org.jspecify.annotations.Nullable;

/**
 * Reads tables with known number of columns, known header labels 
 * and known start position of the table.
 * 
 * <p>It obtains cell values as {@code Cell} object.</p>
 * 
 * <p>The header line is required.
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in a line.</p>
 */
public class CellOneLineHeaderExcelTableReader extends ExcelTableReader<Cell>
    implements IfFormatOneLineHeaderExcelTableReader<Cell>, IfDataTypeCellExcelTableReader {

  private String[] headerLabels;

  /**
   * Constructs a new instance with the sheet name and header labels.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param sheetName See {@link jp.ecuacion.util.excel.table.ExcelTable#sheetName}.
   * @param headerLabels expected header labels
   */
  public CellOneLineHeaderExcelTableReader(String sheetName, String[] headerLabels) {
    super(sheetName);
    this.headerLabels = ObjectsUtil.requireNonNull(headerLabels);
    setTableColumnSize(getHeaderLabels().length);
  }

  /**
   * Constructs a new instance. the obtained value
   *     from an empty cell is {@code null}.
   *
   * <p>{@code tableColumnSize} is not designated
   *     because {@code tableColumnSize} of the table is obviously equal to
   *     the length of the header array.</p>
   *
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber},
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link ExcelTableReader#ExcelTableReader(String, Integer, int, Integer, Integer)}.</p>
   *
   * @deprecated Use the minimal constructor with fluent setters instead.
   */
  @Deprecated
  public CellOneLineHeaderExcelTableReader(String sheetName, String[] headerLabels,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, null);

    this.headerLabels = ObjectsUtil.requireNonNull(headerLabels);

    // Since "Cannot refer to an instance method while explicitly invoking a constructor",
    // First set "null" in "super(...)" and then set the actual value here.
    setTableColumnSize(getHeaderLabels().length);
  }

  @Override
  public String getFarLeftAndTopHeaderLabel() {
    return Objects.requireNonNull(getHeaderLabels()[0]);
  }

  @Override
  public String[] getHeaderLabels() {
    return headerLabels;
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableStartRowNumber(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableStartRowNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableStartColumnNumber(int value) {
    return (CellOneLineHeaderExcelTableReader) super.tableStartColumnNumber(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableRowSize(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableRowSize(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader tableColumnSize(@Nullable Integer value) {
    return (CellOneLineHeaderExcelTableReader) super.tableColumnSize(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    return (CellOneLineHeaderExcelTableReader)
        super.withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public CellOneLineHeaderExcelTableReader withVerticalAndHorizontalOpposite(boolean value) {
    return (CellOneLineHeaderExcelTableReader) super.withVerticalAndHorizontalOpposite(value);
  }
}
