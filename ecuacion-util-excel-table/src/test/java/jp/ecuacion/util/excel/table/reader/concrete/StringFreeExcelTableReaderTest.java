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
import java.util.List;
import java.util.stream.Stream;
import jp.ecuacion.util.excel.enums.NoDataString;
import jp.ecuacion.util.excel.exception.ExcelAppException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

@DisplayName("StringFreeExcelTableReader")
public class StringFreeExcelTableReaderTest {

  /**
   * Creates or reuses a row, then sets a cell.
   * When value is null, creates a BLANK cell. When value is non-null, sets the string value.
   */
  private static void setCell(Sheet sheet, int poiRow, int poiCol, @Nullable String value) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      row = sheet.createRow(poiRow);
    }
    if (value == null) {
      row.createCell(poiCol);
    } else {
      row.createCell(poiCol).setCellValue(value);
    }
  }

  @Nested
  @DisplayName("正常取得")
  class NormalRead {

    @Test
    @DisplayName("通常テーブル（2行×3列、全セル有値）→ 全データを取得できる")
    void normalTable() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "data1-1");
        setCell(sheet, 0, 1, "data1-2");
        setCell(sheet, 0, 2, "data1-3");
        setCell(sheet, 1, 0, "data2-1");
        setCell(sheet, 1, 1, "data2-2");
        setCell(sheet, 1, 2, "data2-3");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, null, null).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("data1-1", "data1-2", "data1-3");
        assertThat(result.get(1)).containsExactly("data2-1", "data2-2", "data2-3");
      }
    }

    @ParameterizedTest(name = "[{index}] noDataString={0} → 空セルが {1} で返る")
    @MethodSource
    @DisplayName("空セルの返り値は noDataString に従う")
    void noDataString(NoDataString noDataString, @Nullable String expected) throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "hello");
        setCell(sheet, 0, 1, null); // BLANK cell
        setCell(sheet, 0, 2, "world");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, 1, 3, noDataString).read(wb);

        assertThat(result.get(0)).containsExactly("hello", expected, "world");
      }
    }

    static Stream<Arguments> noDataString() {
      return Stream.of(
          Arguments.of(NoDataString.NULL, null),
          Arguments.of(NoDataString.EMPTY_STRING, ""));
    }
  }

  @Nested
  @DisplayName("tableRowSize")
  class TableRowSizeTests {

    @Test
    @DisplayName("tableRowSize 指定あり、データ行が指定数より少ない → 不足分は空リスト")
    void fixedRowSizeExceedsData() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 1, 0, "c");
        setCell(sheet, 1, 1, "d");
        // Row 2 not created → empty

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, 3, 2).read(wb);

        assertThat(result).hasSize(3);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).containsExactly("c", "d");
        assertThat(result.get(2)).isEmpty();
      }
    }

    @Test
    @DisplayName("tableRowSize=null → 最初の全空行で読み取り終了")
    void autoRowSizeStopsAtEmptyRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 1, 0, "c");
        setCell(sheet, 1, 1, "d");
        // Row 2 not created → empty row → stops here
        setCell(sheet, 3, 0, "e"); // should NOT be read
        setCell(sheet, 3, 1, "f");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, null, 2).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).containsExactly("c", "d");
      }
    }

    @Test
    @DisplayName("tableRowSize 指定あり、途中に空行 → 空リストとして含まれ打ち切りにならない")
    void emptyRowWithinFixedSize() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        // Row 1 not created → empty
        setCell(sheet, 2, 0, "c");
        setCell(sheet, 2, 1, "d");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, 3, 2).read(wb);

        assertThat(result).hasSize(3);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).isEmpty();
        assertThat(result.get(2)).containsExactly("c", "d");
      }
    }
  }

  @Nested
  @DisplayName("tableColumnSize")
  class TableColumnSizeTests {

    @Test
    @DisplayName("tableColumnSize 指定あり → 指定列数のみ取得、それ以降は無視")
    void fixedColumnSize() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 0, 2, "c"); // col 2 beyond tableColumnSize=2 → ignored
        setCell(sheet, 0, 3, "d");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, 1, 2).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("a", "b");
      }
    }

    @Test
    @DisplayName("tableColumnSize=null → 最初の行の非空セル連続数で自動決定")
    void autoColumnSize() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 0, 2, "c");
        // Col 3 not created → auto column size stops at 3
        setCell(sheet, 0, 4, "extra"); // beyond the break → not read
        setCell(sheet, 1, 0, "d");
        setCell(sheet, 1, 1, "e");
        setCell(sheet, 1, 2, "f");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, null, null).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("a", "b", "c");
        assertThat(result.get(1)).containsExactly("d", "e", "f");
      }
    }
  }

  @Nested
  @DisplayName("開始位置")
  class StartPosition {

    @Test
    @DisplayName("tableStartRowNumber=3, tableStartColumnNumber=2 → 指定位置からデータ取得")
    void offsetPosition() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "unrelated"); // outside table
        // table starts at poi row=2, poi col=1 (= tableStartRow=3, tableStartCol=2)
        setCell(sheet, 2, 1, "a");
        setCell(sheet, 2, 2, "b");
        setCell(sheet, 3, 1, "c");
        setCell(sheet, 3, 2, "d");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 3, 2, null, null).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).containsExactly("c", "d");
      }
    }
  }

  @Nested
  @DisplayName("isVerticalAndHorizontalOpposite")
  class VerticalTable {

    @Test
    @DisplayName("isVerticalAndHorizontalOpposite=true → 縦向きテーブルを行列入れ替えて取得")
    void verticalTable() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Physical layout: each physical column = one logical data row
        // Col 0 → logical row 0, Col 1 → logical row 1
        setCell(sheet, 0, 0, "r1c0");
        setCell(sheet, 0, 1, "r2c0");
        setCell(sheet, 1, 0, "r1c1");
        setCell(sheet, 1, 1, "r2c1");
        // Col 2 not created → terminates

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1", 1, 1, null, null)
                .isVerticalAndHorizontalOpposite(true).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("r1c0", "r1c1");
        assertThat(result.get(1)).containsExactly("r2c0", "r2c1");
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
        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("NotExist", 1, 1, null, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.SheetNotExist.message");
      }
    }

    @Test
    @DisplayName("テーブル開始位置にデータがない → ExcelAppException（ColumnSizeIsZero）")
    void columnSizeIsZero() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1"); // empty sheet
        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("Sheet1", 1, 1, null, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.reader.ColumnSizeIsZero.message");
      }
    }
  }
}
