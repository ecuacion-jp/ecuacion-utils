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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("TypedOneLineHeaderExcelTableReader / TypedHeaderExcelTableReader")
public class TypedHeaderExcelTableReaderTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  // --- helpers ---

  private static Row getOrCreateRow(Sheet sheet, int poiRow) {
    Row row = sheet.getRow(poiRow);
    return row == null ? sheet.createRow(poiRow) : row;
  }

  private static void setStringCell(Sheet sheet, int poiRow, int poiCol, String value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setNumericCell(Sheet sheet, int poiRow, int poiCol, double value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setBooleanCell(Sheet sheet, int poiRow, int poiCol, boolean value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setBlankCell(Sheet sheet, int poiRow, int poiCol) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol);
  }

  private static void setErrorCell(Sheet sheet, int poiRow, int poiCol, FormulaError error) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellErrorValue(error.getCode());
  }

  private static void setDateFormattedCell(Workbook wb, Sheet sheet, int poiRow, int poiCol,
      LocalDateTime value, String formatPattern) {
    Cell cell = getOrCreateRow(sheet, poiRow).createCell(poiCol);
    CellStyle style = wb.createCellStyle();
    style.setDataFormat(wb.createDataFormat().getFormat(formatPattern));
    cell.setCellStyle(style);
    cell.setCellValue(value);
  }

  private Path writeTempExcel(Workbook wb) throws Exception {
    Path file = tempDir.resolve("test.xlsx");
    try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
      wb.write(fos);
    }
    return file;
  }

  @Nullable
  private Object readSingleCell(Path file) throws Exception {
    var reader = new TypedOneLineHeaderExcelTableReader("Sheet1", new String[] {"value"})
        .tableStartRowNumber(1);
    List<List<Object>> result = reader.read(file.toString());
    return result.get(0).get(0);
  }

  /**
   * Reads column "value" of the single data row, alongside a "marker" column that always holds
   * a non-empty value — needed when "value" itself is blank/empty, since a row whose every
   * cell is empty is treated as the end of the table and skipped.
   */
  @Nullable
  private Object readValueWithMarker(Path file) throws Exception {
    var reader = new TypedOneLineHeaderExcelTableReader("Sheet1", new String[] {"value", "marker"})
        .tableStartRowNumber(1);
    List<List<Object>> result = reader.read(file.toString());
    return result.get(0).get(0);
  }

  @Nested
  @DisplayName("セルの型変換")
  class CellTypeConversion {

    @Test
    @DisplayName("文字列セル → String")
    void stringCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setStringCell(sheet, 1, 0, "Alice");
        Path file = writeTempExcel(wb);

        assertThat(readSingleCell(file)).isEqualTo("Alice");
      }
    }

    @Test
    @DisplayName("空文字列セル → null")
    void emptyStringCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setStringCell(sheet, 0, 1, "marker");
        setStringCell(sheet, 1, 0, "");
        setStringCell(sheet, 1, 1, "x");
        Path file = writeTempExcel(wb);

        assertThat(readValueWithMarker(file)).isNull();
      }
    }

    @Test
    @DisplayName("数値セル（日付書式なし） → Double")
    void numericCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setNumericCell(sheet, 1, 0, 25.0);
        Path file = writeTempExcel(wb);

        assertThat(readSingleCell(file)).isEqualTo(25.0);
      }
    }

    @Test
    @DisplayName("日付書式の数値セル（時刻が0時0分） → LocalDate")
    void dateFormattedCellAtMidnight() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setDateFormattedCell(wb, sheet, 1, 0, LocalDateTime.of(2026, 1, 15, 0, 0), "yyyy-mm-dd");
        Path file = writeTempExcel(wb);

        assertThat(readSingleCell(file)).isEqualTo(LocalDate.of(2026, 1, 15));
      }
    }

    @Test
    @DisplayName("日付書式の数値セル（時刻あり） → LocalDateTime")
    void dateFormattedCellWithTime() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setDateFormattedCell(wb, sheet, 1, 0, LocalDateTime.of(2026, 1, 15, 9, 30),
            "yyyy-mm-dd hh:mm:ss");
        Path file = writeTempExcel(wb);

        assertThat(readSingleCell(file)).isEqualTo(LocalDateTime.of(2026, 1, 15, 9, 30));
      }
    }

    @Test
    @DisplayName("真偽値セル → Boolean")
    void booleanCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setBooleanCell(sheet, 1, 0, true);
        Path file = writeTempExcel(wb);

        assertThat(readSingleCell(file)).isEqualTo(Boolean.TRUE);
      }
    }

    @Test
    @DisplayName("空白セル → null")
    void blankCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setStringCell(sheet, 0, 1, "marker");
        setBlankCell(sheet, 1, 0);
        setStringCell(sheet, 1, 1, "x");
        Path file = writeTempExcel(wb);

        assertThat(readValueWithMarker(file)).isNull();
      }
    }

    @Test
    @DisplayName("エラーセル → ExcelTableException")
    void errorCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "value");
        setErrorCell(sheet, 1, 0, FormulaError.DIV0);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableReader("Sheet1", new String[] {"value"})
            .tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(file.toString()))
            .isInstanceOf(ExcelTableException.class);
      }
    }
  }

  @Nested
  @DisplayName("正常取得")
  class NormalRead {

    @Test
    @DisplayName("複数列・複数行 → ヘッダーは除外され型変換済みの値で返る")
    void multiColumnMultiRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "score");
        setStringCell(sheet, 0, 2, "active");
        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, 92.5);
        setBooleanCell(sheet, 1, 2, true);
        setStringCell(sheet, 2, 0, "Bob");
        setNumericCell(sheet, 2, 1, 78.0);
        setBooleanCell(sheet, 2, 2, false);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableReader("Sheet1",
            new String[] {"name", "score", "active"}).tableStartRowNumber(1);
        List<List<Object>> result = reader.read(file.toString());

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("Alice", 92.5, true);
        assertThat(result.get(1)).containsExactly("Bob", 78.0, false);
      }
    }
  }
}
