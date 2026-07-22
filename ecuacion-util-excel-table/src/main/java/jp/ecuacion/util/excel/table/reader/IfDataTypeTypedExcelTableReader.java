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
package jp.ecuacion.util.excel.table.reader;

import java.time.LocalDateTime;
import java.time.LocalTime;
import jp.ecuacion.util.excel.exception.CellContainsErrorException;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.table.IfDataTypeTypedExcelTable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.jspecify.annotations.Nullable;

/**
 * Provides the excel table reader interface
 *     with object type obtained from the excel data being a native Java type.
 *
 * <p>Cell values are converted as follows:</p>
 * <ul>
 *   <li>Blank cells → {@code null}</li>
 *   <li>String cells → {@link String} ({@code null} when empty)</li>
 *   <li>Numeric cells with date format → {@link java.time.LocalDate} when no time component,
 *       {@link LocalDateTime} otherwise</li>
 *   <li>Numeric cells without date format → {@link Double}</li>
 *   <li>Boolean cells → {@link Boolean}</li>
 * </ul>
 */
public interface IfDataTypeTypedExcelTableReader
    extends IfDataTypeTypedExcelTable, IfExcelTableReader<Object> {

  @Override
  public default @Nullable Object getCellData(Cell cell, int columnNumber)
      throws ExcelTableException {
    CellType cellType = cell.getCellType();
    if (cellType == CellType.FORMULA) {
      cellType = cell.getCachedFormulaResultType();
    }

    if (cellType == CellType.BLANK) {
      return null;
    } else if (cellType == CellType.STRING) {
      String v = cell.getStringCellValue();
      return v.isEmpty() ? null : v;
    } else if (cellType == CellType.NUMERIC) {
      if (DateUtil.isCellDateFormatted(cell)) {
        LocalDateTime ldt = cell.getLocalDateTimeCellValue();
        return ldt.toLocalTime().equals(LocalTime.MIDNIGHT) ? ldt.toLocalDate() : ldt;
      }
      return cell.getNumericCellValue();
    } else if (cellType == CellType.BOOLEAN) {
      return cell.getBooleanCellValue();
    } else if (cellType == CellType.ERROR) {
      throw new CellContainsErrorException(cell.getRow().getSheet().getSheetName(),
          cell.getAddress().formatAsString(), null);
    } else {
      throw new RuntimeException("cell type not found. cellType: " + cellType);
    }
  }

  @Override
  public default boolean isCellDataEmpty(@Nullable Object cellData) {
    if (cellData == null) {
      return true;
    }
    if (cellData instanceof String s) {
      return s.isEmpty();
    }
    return false;
  }
}
