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
package jp.ecuacion.util.excel.table.writer;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Map;
import jp.ecuacion.util.excel.table.IfDataTypeTypedExcelTable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Provides the excel table writer interface
 *     with object type written to the excel data being a native Java type.
 *
 * <p>{@link LocalDate} and {@link LocalDateTime} values are written as date-formatted cells
 *     (not strings): the destination cell keeps its existing date format if it already has one,
 *     otherwise a date-formatted {@link CellStyle} is created and applied so the value is
 *     displayed as a date in Excel. {@link Number} values are written as numeric cells,
 *     {@link Boolean} values as boolean cells, and any other value is written via
 *     {@link Object#toString()}.</p>
 */
public interface IfDataTypeTypedExcelTableWriter
    extends IfDataTypeTypedExcelTable, IfExcelTableWriter<Object> {

  @Override
  public default void writeToCell(int columnNumberFromZero, @Nullable Object sourceCellData,
      Cell destCell) {
    if (sourceCellData == null) {
      return;
    }

    if (sourceCellData instanceof LocalDate date) {
      destCell.setCellValue(date);
      ensureDateCellStyle(destCell, getDateFormat());

    } else if (sourceCellData instanceof LocalDateTime dateTime) {
      destCell.setCellValue(dateTime);
      ensureDateCellStyle(destCell, getDateTimeFormat());

    } else if (sourceCellData instanceof Number number) {
      destCell.setCellValue(number.doubleValue());

    } else if (sourceCellData instanceof Boolean bool) {
      destCell.setCellValue(bool);

    } else {
      destCell.setCellValue(sourceCellData.toString());
    }
  }

  private void ensureDateCellStyle(Cell destCell, String formatPattern) {
    if (DateUtil.isCellDateFormatted(destCell)) {
      return;
    }

    Map<String, CellStyle> styleMap = getDateCellStyleMap();
    CellStyle style = styleMap.get(formatPattern);
    if (style == null) {
      Workbook workbook = destCell.getRow().getSheet().getWorkbook();
      style = workbook.createCellStyle();
      style.cloneStyleFrom(destCell.getCellStyle());
      style.setDataFormat(workbook.createDataFormat().getFormat(formatPattern));
      styleMap.put(formatPattern, style);
    }

    destCell.setCellStyle(style);
  }

  /**
   * Returns the Excel number-format pattern applied to {@link LocalDate} values when the
   *     destination cell does not already have a date format.
   *
   * <p>Defaults to {@code "yyyy-mm-dd"}.</p>
   *
   * @return the format pattern
   */
  public String getDateFormat();

  /**
   * Returns the Excel number-format pattern applied to {@link LocalDateTime} values when the
   *     destination cell does not already have a date format.
   *
   * <p>Defaults to {@code "yyyy-mm-dd hh:mm:ss"}.</p>
   *
   * @return the format pattern
   */
  public String getDateTimeFormat();

  /**
   * Returns the cache of {@link CellStyle}s created for date formatting, keyed by format pattern.
   *
   * <p>The number of {@code CellStyle}s in an excel file has a limit (64,000), so styles created
   *     for date formatting are cached here and reused across cells.</p>
   *
   * @return the style cache
   */
  public Map<String, CellStyle> getDateCellStyleMap();
}
