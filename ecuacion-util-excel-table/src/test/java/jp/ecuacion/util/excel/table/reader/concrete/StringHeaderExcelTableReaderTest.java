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
import jp.ecuacion.util.excel.exception.FarLeftHeaderLabelNotFoundException;
import jp.ecuacion.util.excel.exception.HeaderCellIsBlankException;
import jp.ecuacion.util.excel.exception.NumberOfTableHeadersDifferException;
import jp.ecuacion.util.excel.exception.TableHeaderTitleWrongException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

// Common ExcelTableReader base-class behaviors (tableRowSize, start position,
// withVerticalAndHorizontalOpposite, SheetNotExist, etc.) are covered by StringFreeExcelTableReaderTest.
// This test class covers only behaviors specific to this reader.
@DisplayName("StringOneLineHeaderExcelTableReader / StringHeaderExcelTableReader"
    + " Note: see StringFreeExcelTableReaderTest for common base-class behavior")
public class StringHeaderExcelTableReaderTest {

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
  @DisplayName("Normal read")
  class NormalRead {

    @Test
    @DisplayName("header row + data rows → header excluded, only data returned")
    void normalTable() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "header1");
        setCell(sheet, 0, 1, "header2");
        setCell(sheet, 0, 2, "header3");
        setCell(sheet, 1, 0, "data1-1");
        setCell(sheet, 1, 1, "data1-2");
        setCell(sheet, 1, 2, "data1-3");
        setCell(sheet, 2, 0, "data2-1");
        setCell(sheet, 2, 1, "data2-2");
        setCell(sheet, 2, 2, "data2-3");

        List<List<String>> result = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2", "header3"})
            .tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("data1-1", "data1-2", "data1-3");
        assertThat(result.get(1)).containsExactly("data2-1", "data2-2", "data2-3");
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=null → auto-detects row position by header label")
    void autoDetectStartRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "unrelated");
        setCell(sheet, 1, 0, "another");
        // Table starts at row 2
        setCell(sheet, 2, 0, "header1");
        setCell(sheet, 2, 1, "header2");
        setCell(sheet, 3, 0, "data1");
        setCell(sheet, 3, 1, "data2");

        List<List<String>> result = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2"}).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("data1", "data2");
      }
    }
  }

  @Nested
  @DisplayName("Header validation")
  class HeaderValidation {

    @Test
    @DisplayName("Excel column count > expected column count, ignores=false"
        + " → NumberOfTableHeadersDifferException")
    void tooManyColumnsIgnoresFalse() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2");
        setCell(sheet, 0, 2, "h3");
        setCell(sheet, 0, 3, "extra"); // 4 columns, expected 3

        StringOneLineHeaderExcelTableReader reader = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}).tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(NumberOfTableHeadersDifferException.class);
      }
    }

    @Test
    @DisplayName("Excel column count > expected column count, ignores=true → only expected columns of data are retrieved")
    void tooManyColumnsIgnoresTrue() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2");
        setCell(sheet, 0, 2, "h3");
        setCell(sheet, 0, 3, "extra"); // 4 columns, expected 3
        setCell(sheet, 1, 0, "d1");
        setCell(sheet, 1, 1, "d2");
        setCell(sheet, 1, 2, "d3");
        setCell(sheet, 1, 3, "d4");

        List<List<String>> result = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}).tableStartRowNumber(1)
            .withIgnoresAdditionalColumnsOfHeaderData(true).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("d1", "d2", "d3");
      }
    }

    @ParameterizedTest(name = "[{index}] ignores={0} → NumberOfTableHeadersDifferException")
    @MethodSource
    @DisplayName("Excel column count < expected column count"
        + " → NumberOfTableHeadersDifferException regardless of ignores setting")
    void tooFewColumns(boolean ignores) throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2"); // 2 columns, expected 3

        var reader = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}).tableStartRowNumber(1)
            .withIgnoresAdditionalColumnsOfHeaderData(ignores);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(NumberOfTableHeadersDifferException.class);
      }
    }

    static @Nullable Stream<@Nullable Arguments> tooFewColumns() {
      return Stream.of(
          Arguments.of(false),
          Arguments.of(true));
    }

    @Test
    @DisplayName("header label text mismatch → TableHeaderTitleWrongException")
    void labelMismatch() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "WRONG"); // expected "h2"

        StringOneLineHeaderExcelTableReader reader = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2"}).tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(TableHeaderTitleWrongException.class);
      }
    }
  }

  @Nested
  @DisplayName("Error cases")
  class ErrorCases {

    @Test
    @DisplayName("tableStartRowNumber=null, header label not found"
        + " → FarLeftHeaderLabelNotFoundException")
    void headerLabelNotFound() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "unrelated");

        StringOneLineHeaderExcelTableReader reader = new StringOneLineHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2"});
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(FarLeftHeaderLabelNotFoundException.class);
      }
    }
  }

  @Nested
  @DisplayName("Multi-row header")
  class MultiLineHeader {

    @Test
    @DisplayName("two-row header → header excluded, only data returned")
    void twoRowHeader() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // header row 0: group labels
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "PersonalInfo");
        setCell(sheet, 0, 2, "PersonalInfo");
        // header row 1: column labels
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");
        // data
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}})
            .tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("all header rows are validated (row 1 mismatch"
        + " → TableHeaderTitleWrongException)")
    void firstRowMismatch() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "WRONG"); // expected "PersonalInfo"
        setCell(sheet, 0, 2, "PersonalInfo");
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");

        var reader = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}})
            .tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(TableHeaderTitleWrongException.class);
      }
    }

    @Test
    @DisplayName("horizontally merged cell is expanded and validated correctly")
    void horizontalMergedCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // "PersonalInfo" merged over cols 1-2
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "PersonalInfo"); // master cell
        // col 2 is empty because it's part of the merge
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}})
            .tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("vertically merged cell (# column) is expanded and validated correctly")
    void verticalMergedCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // "#" merged vertically over rows 0-1
        setCell(sheet, 0, 0, "#"); // master
        // row 1 col 0 empty (part of vertical merge)
        setCell(sheet, 0, 1, "PersonalInfo");
        setCell(sheet, 0, 2, "PersonalInfo");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}})
            .tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("blank header cell with no merge → HeaderCellIsBlankException")
    void blankNonMergedHeaderCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "PersonalInfo");
        // col 2 is blank but NOT part of any merge
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");

        var reader = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}})
            .tableStartRowNumber(1);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(HeaderCellIsBlankException.class);
      }
    }
  }
}
