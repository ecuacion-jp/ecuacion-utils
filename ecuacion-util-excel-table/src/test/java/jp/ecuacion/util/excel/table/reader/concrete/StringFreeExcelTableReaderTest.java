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
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.stream.Stream;
import jp.ecuacion.util.excel.enums.NoDataString;
import jp.ecuacion.util.excel.exception.ColumnSizeIsZeroException;
import jp.ecuacion.util.excel.exception.SheetNotExistException;
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
  @DisplayName("normal read")
  class NormalRead {

    @Test
    @DisplayName("normal table (2 rows x 3 columns, all cells have values) → all data can be retrieved")
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
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("data1-1", "data1-2", "data1-3");
        assertThat(result.get(1)).containsExactly("data2-1", "data2-2", "data2-3");
      }
    }

    @ParameterizedTest(name = "[{index}] noDataString={0} → empty cell returns as {1}")
    @MethodSource
    @DisplayName("the return value for an empty cell follows noDataString")
    void noDataString(NoDataString noDataString, @Nullable String expected) throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "hello");
        setCell(sheet, 0, 1, null); // BLANK cell
        setCell(sheet, 0, 2, "world");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableRowSize(1)
                .tableColumnSize(3).noDataString(noDataString).read(wb);

        assertThat(result.get(0)).containsExactly("hello", expected, "world");
      }
    }

    static @Nullable Stream<@Nullable Arguments> noDataString() {
      return Stream.of(
          Arguments.of(NoDataString.NULL, null),
          Arguments.of(NoDataString.EMPTY_STRING, ""));
    }
  }

  @Nested
  @DisplayName("tableRowSize")
  class TableRowSizeTests {

    @Test
    @DisplayName("tableRowSize is specified, and there are fewer data rows than specified → the shortfall is an empty list")
    void fixedRowSizeExceedsData() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 1, 0, "c");
        setCell(sheet, 1, 1, "d");
        // Row 2 not created → empty

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableRowSize(3)
                .tableColumnSize(2).read(wb);

        assertThat(result).hasSize(3);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).containsExactly("c", "d");
        assertThat(result.get(2)).isEmpty();
      }
    }

    @Test
    @DisplayName("tableRowSize=null → reading ends at the first fully empty row")
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
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableColumnSize(2)
                .read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("a", "b");
        assertThat(result.get(1)).containsExactly("c", "d");
      }
    }

    @Test
    @DisplayName("tableRowSize is specified, with an empty row in the middle → included as an empty list, not truncated")
    void emptyRowWithinFixedSize() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        // Row 1 not created → empty
        setCell(sheet, 2, 0, "c");
        setCell(sheet, 2, 1, "d");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableRowSize(3)
                .tableColumnSize(2).read(wb);

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
    @DisplayName("tableColumnSize is specified → only the specified number of columns is retrieved, the rest is ignored")
    void fixedColumnSize() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 0, 2, "c"); // col 2 beyond tableColumnSize=2 → ignored
        setCell(sheet, 0, 3, "d");

        List<List<String>> result =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableRowSize(1)
                .tableColumnSize(2).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("a", "b");
      }
    }

    @Test
    @DisplayName("tableColumnSize=null → automatically determined by the run of non-empty cells in the first row")
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
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("a", "b", "c");
        assertThat(result.get(1)).containsExactly("d", "e", "f");
      }
    }
  }

  @Nested
  @DisplayName("start position")
  class StartPosition {

    @Test
    @DisplayName("tableStartRowNumber=3, tableStartColumnNumber=2 → retrieves data from the specified position")
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
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(3)
                .tableStartColumnNumber(2).read(wb);

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
    @DisplayName("isVerticalAndHorizontalOpposite=true → retrieves a vertical table with rows and columns swapped")
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
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1)
                .withVerticalAndHorizontalOpposite(true).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("r1c0", "r1c1");
        assertThat(result.get(1)).containsExactly("r2c0", "r2c1");
      }
    }
  }

  @Nested
  @DisplayName("error cases")
  class ErrorCases {

    @Test
    @DisplayName("nonexistent sheet name → SheetNotExistException")
    void sheetNotExist() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1");
        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("NotExist").tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(SheetNotExistException.class);
      }
    }

    @Test
    @DisplayName("no data at the table start position → ColumnSizeIsZeroException")
    void columnSizeIsZero() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        wb.createSheet("Sheet1"); // empty sheet
        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ColumnSizeIsZeroException.class);
      }
    }
  }

  @Nested
  @DisplayName("getIterable")
  class IterableReaderTests {

    @Test
    @DisplayName("getIterable(Workbook) → all rows can be retrieved in order with a for-each loop")
    void iterateAllRows() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");
        setCell(sheet, 0, 1, "b");
        setCell(sheet, 1, 0, "c");
        setCell(sheet, 1, 1, "d");

        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableColumnSize(2);
        List<List<String>> collected = new ArrayList<>();
        try (var iterable = reader.getIterable(wb)) {
          for (List<String> row : iterable) {
            collected.add(row);
          }
        }

        assertThat(collected).hasSize(2);
        assertThat(collected.get(0)).containsExactly("a", "b");
        assertThat(collected.get(1)).containsExactly("c", "d");
      }
    }

    @Test
    @DisplayName("after all rows are consumed, hasNext() becomes false and calling next() throws NoSuchElementException")
    void exhaustedIterator() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "a");

        StringFreeExcelTableReader reader =
            new StringFreeExcelTableReader("Sheet1").tableStartRowNumber(1).tableColumnSize(1);
        try (var iterable = reader.getIterable(wb)) {
          var iterator = iterable.iterator();
          iterator.next(); // consume the only row → sets hasNext=false

          assertThat(iterator.hasNext()).isFalse();
          assertThatThrownBy(iterator::next).isInstanceOf(NoSuchElementException.class);
        }
      }
    }
  }
}
