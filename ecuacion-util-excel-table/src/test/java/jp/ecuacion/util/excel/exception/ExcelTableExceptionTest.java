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
package jp.ecuacion.util.excel.exception;

import static org.assertj.core.api.Assertions.assertThat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("ExcelTableException")
public class ExcelTableExceptionTest {

  @Nested
  @DisplayName("cell(Cell)")
  class CellMethod {

    @Test
    @DisplayName("cell() sets cell, sheet, and workbook from the given Cell object")
    void cellSetsCellSheetAndWorkbook() {
      Workbook wb = new XSSFWorkbook();
      Sheet sheet = wb.createSheet("Sheet1");
      Row row = sheet.createRow(0);
      Cell cell = row.createCell(0);
      cell.setCellValue("test");

      ExcelTableException ex = new ExcelTableException("msg.id");
      ex.cell(cell);

      assertThat(ex.getCell()).isSameAs(cell);
      assertThat(ex.getSheet()).isSameAs(sheet);
      assertThat(ex.getWorkbook()).isSameAs(wb);
    }

    @Test
    @DisplayName("cell() returns self for method chaining")
    @SuppressWarnings("resource")
    void cellReturnsSelf() {
      Workbook wb = new XSSFWorkbook();
      Sheet sheet = wb.createSheet("Sheet1");
      Row row = sheet.createRow(0);
      Cell cell = row.createCell(0);

      ExcelTableException ex = new ExcelTableException("msg.id");
      ExcelTableException result = ex.cell(cell);

      assertThat(result).isSameAs(ex);
    }
  }
}
