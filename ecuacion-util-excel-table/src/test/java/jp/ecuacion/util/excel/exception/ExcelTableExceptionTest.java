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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("ExcelTableException")
public class ExcelTableExceptionTest {

  @Nested
  @DisplayName("cell position format (jp.ecuacion.util.excel.cell-format-r1c1)")
  class CellPositionFormat {

    private static final String KEY = "jp.ecuacion.util.excel.cell-format-r1c1";

    @AfterEach
    void clearProperty() {
      System.clearProperty(KEY);
    }

    @Test
    @DisplayName("defaults to A1-style address when the property is not set")
    void defaultsToA1Format() {
      HeaderCellIsBlankException ex = new HeaderCellIsBlankException("Sheet1", 1, 2);

      assertThat(ex.getViolations().getBusinessViolations().get(0).toString())
          .contains("target cell: B1");
    }

    @Test
    @DisplayName("resolves as row/column numbers when the property is \"true\"")
    void resolvesAsRowColumnWhenPropertyIsTrue() {
      System.setProperty(KEY, "true");

      HeaderCellIsBlankException ex = new HeaderCellIsBlankException("Sheet1", 1, 2);

      assertThat(ex.getViolations().getBusinessViolations().get(0).toString())
          .contains("row number: 1, column number: 2");
    }

    @Test
    @DisplayName("converts 1-based row/column to the expected A1 address")
    void convertsRowColumnToExpectedA1Address() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // 0-based row 2 / column 26 == 1-based row 3 / column 27 (AA)
        Cell cell = sheet.createRow(2).createCell(26);

        CellContainsErrorException ex = new CellContainsErrorException(cell);

        assertThat(ex.getViolations().getBusinessViolations().get(0).toString())
            .contains("target cell: AA3");
      }
    }
  }
}
