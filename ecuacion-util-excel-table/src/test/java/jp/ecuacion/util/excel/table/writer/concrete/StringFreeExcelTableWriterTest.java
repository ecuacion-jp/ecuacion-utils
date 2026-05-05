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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.util.List;
import jp.ecuacion.util.excel.exception.ExcelAppException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("StringFreeExcelTableWriter")
public class StringFreeExcelTableWriterTest {

  @Nested
  @DisplayName("正常書き込み")
  class NormalWrite {

    @Test
    @DisplayName("通常書き込み（2行×3列）→ 指定セルに String 値が書き込まれる")
    void normalWrite() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1");
        List<List<String>> data = List.of(
            List.of("a", "b", "c"),
            List.of("d", "e", "f"));

        new StringFreeExcelTableWriter("Sheet1", 1, 1).write(wb, data);

        Sheet sheet = wb.getSheet("Sheet1");
        assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("a");
        assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("b");
        assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("c");
        assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("d");
        assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("e");
        assertThat(sheet.getRow(1).getCell(2).getStringCellValue()).isEqualTo("f");
      }
    }
  }

  @Nested
  @DisplayName("開始位置")
  class StartPosition {

    @Test
    @DisplayName("tableStartRowNumber=3, tableStartColumnNumber=2 → 指定位置から書き込まれる")
    void offsetPosition() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1");
        List<List<String>> data = List.of(List.of("a", "b"));

        // tableStartRow=3 → poi row=2, tableStartCol=2 → poi col=1
        new StringFreeExcelTableWriter("Sheet1", 3, 2).write(wb, data);

        Sheet sheet = wb.getSheet("Sheet1");
        assertThat(sheet.getRow(2).getCell(1).getStringCellValue()).isEqualTo("a");
        assertThat(sheet.getRow(2).getCell(2).getStringCellValue()).isEqualTo("b");
        assertThat(sheet.getRow(0)).isNull(); // rows before start are untouched
      }
    }
  }

  @Nested
  @DisplayName("isVerticalAndHorizontalOpposite")
  class VerticalTable {

    @Test
    @DisplayName("isVerticalAndHorizontalOpposite=true → 縦向きに書き込まれる")
    void verticalWrite() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1");
        // Logical: row 0 = ["r1c0","r1c1"], row 1 = ["r2c0","r2c1"]
        // Physical result: col 0 = first row, col 1 = second row
        List<List<String>> data = List.of(
            List.of("r1c0", "r1c1"),
            List.of("r2c0", "r2c1"));

        new StringFreeExcelTableWriter("Sheet1", 1, 1)
            .isVerticalAndHorizontalOpposite(true).write(wb, data);

        Sheet sheet = wb.getSheet("Sheet1");
        assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("r1c0");
        assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("r1c1");
        assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("r2c0");
        assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("r2c1");
      }
    }
  }

  @Nested
  @DisplayName("異常系")
  class ErrorCases {

    @Test
    @DisplayName("存在しないシート名 → ExcelAppException（SheetNotExist）")
    void sheetNotExist() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1");
        StringFreeExcelTableWriter writer =
            new StringFreeExcelTableWriter("NotExist", 1, 1);
        assertThatThrownBy(() -> writer.write(wb, List.of(List.of("a"))))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.SheetNotExist.message");
      }
    }
  }
}
