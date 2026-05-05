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
import jp.ecuacion.util.excel.exception.ExcelAppException;
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

// 基底クラス（ExcelTableReader）共通の振る舞い（tableRowSize, 開始位置, isVerticalAndHorizontalOpposite,
// SheetNotExist 等）は StringFreeExcelTableReaderTest でカバー済み。固有の振る舞いのみを扱う。
@DisplayName("StringHeaderExcelTableReader"
    + " ※基底クラス共通の振る舞いは StringFreeExcelTableReaderTest 参照")
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
  @DisplayName("正常取得")
  class NormalRead {

    @Test
    @DisplayName("ヘッダー行 + データ行 → ヘッダーは除外されデータのみ返る")
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

        List<List<String>> result = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2", "header3"}, 1, 1, null).read(wb);

        assertThat(result).hasSize(2);
        assertThat(result.get(0)).containsExactly("data1-1", "data1-2", "data1-3");
        assertThat(result.get(1)).containsExactly("data2-1", "data2-2", "data2-3");
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=null → ヘッダーラベルで行位置を自動検索して取得")
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

        List<List<String>> result = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2"}, null, 1, null).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("data1", "data2");
      }
    }
  }

  @Nested
  @DisplayName("ヘッダー検証")
  class HeaderValidation {

    @Test
    @DisplayName("Excel の列数 > 期待列数、ignores=false → ExcelAppException")
    void tooManyColumnsIgnoresFalse() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2");
        setCell(sheet, 0, 2, "h3");
        setCell(sheet, 0, 3, "extra"); // 4 columns, expected 3

        StringHeaderExcelTableReader reader = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}, 1, 1, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.NumberOfTableHeadersDiffer.message");
      }
    }

    @Test
    @DisplayName("Excel の列数 > 期待列数、ignores=true → データは期待列数分のみ取得")
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

        List<List<String>> result = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}, 1, 1, null)
            .ignoresAdditionalColumnsOfHeaderData(true).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("d1", "d2", "d3");
      }
    }

    @ParameterizedTest(name = "[{index}] ignores={0} → ExcelAppException")
    @MethodSource
    @DisplayName("Excel の列数 < 期待列数 → ignores 設定に関係なく ExcelAppException")
    void tooFewColumns(boolean ignores) throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2"); // 2 columns, expected 3

        var reader = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2", "h3"}, 1, 1, null)
            .ignoresAdditionalColumnsOfHeaderData(ignores);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.NumberOfTableHeadersDiffer.message");
      }
    }

    static Stream<Arguments> tooFewColumns() {
      return Stream.of(
          Arguments.of(false),
          Arguments.of(true));
    }

    @Test
    @DisplayName("ヘッダーラベルのテキスト不一致 → ExcelAppException（TableHeaderTitleWrong）")
    void labelMismatch() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "WRONG"); // expected "h2"

        StringHeaderExcelTableReader reader = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"h1", "h2"}, 1, 1, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.TableHeaderTitleWrong.message");
      }
    }
  }

  @Nested
  @DisplayName("異常系")
  class ErrorCases {

    @Test
    @DisplayName("tableStartRowNumber=null、ヘッダーラベルが見つからない → ExcelAppException")
    void headerLabelNotFound() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "unrelated");

        StringHeaderExcelTableReader reader = new StringHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2"}, null, 1, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo(
                "jp.ecuacion.util.excel.reader.FarLeftHeaderLabelNotFound.message");
      }
    }
  }

  @Nested
  @DisplayName("複数行ヘッダー")
  class MultiLineHeader {

    @Test
    @DisplayName("2行ヘッダー → ヘッダーを除外してデータのみ返る")
    void twoRowHeader() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // header row 0: group labels
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "個人情報");
        // header row 1: column labels
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        // data
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("全ヘッダー行が検証される（1行目不一致 → ExcelAppException）")
    void firstRowMismatch() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "WRONG"); // expected "個人情報"
        setCell(sheet, 0, 2, "個人情報");
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");

        var reader = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.TableHeaderTitleWrong.message");
      }
    }

    @Test
    @DisplayName("横結合セルが展開されて正しく検証される")
    void horizontalMergedCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // "個人情報" merged over cols 1-2
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報"); // master cell
        // col 2 is empty because it's part of the merge
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("縦結合セル（# 列）が展開されて正しく検証される")
    void verticalMergedCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // "#" merged vertically over rows 0-1
        setCell(sheet, 0, 0, "#"); // master
        // row 1 col 0 empty (part of vertical merge)
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "個人情報");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));

        List<List<String>> result = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null).read(wb);

        assertThat(result).hasSize(1);
        assertThat(result.get(0)).containsExactly("1", "Alice", "25");
      }
    }

    @Test
    @DisplayName("結合なし空欄ヘッダーセル → ExcelAppException")
    void blankNonMergedHeaderCell() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報");
        // col 2 is blank but NOT part of any merge
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");

        var reader = new StringHeaderExcelTableReader("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null);
        assertThatThrownBy(() -> reader.read(wb))
            .isInstanceOf(ExcelAppException.class)
            .extracting(e -> ((ExcelAppException) e).getMessageId())
            .isEqualTo("jp.ecuacion.util.excel.reader.HeaderCellIsBlank.message");
      }
    }
  }
}
