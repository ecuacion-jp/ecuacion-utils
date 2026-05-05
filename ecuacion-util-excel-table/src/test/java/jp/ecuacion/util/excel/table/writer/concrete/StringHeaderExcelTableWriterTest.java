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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.jspecify.annotations.Nullable;

// 基底クラス共通の振る舞い（書き込み, 開始位置, isVerticalAndHorizontalOpposite, SheetNotExist 等）は
// StringFreeExcelTableWriterTest でカバー済み。
@DisplayName("StringHeaderExcelTableWriter ※基底クラス共通の振る舞いは StringFreeExcelTableWriterTest 参照")
public class StringHeaderExcelTableWriterTest {

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

  private static @Nullable String getCellValue(Sheet sheet, int poiRow, int poiCol) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      return null;
    }
    var cell = row.getCell(poiCol);
    return cell == null ? null : cell.getStringCellValue();
  }

  @Nested
  @DisplayName("ヘッダー行の検証と書き込み")
  class HeaderWrite {

    @Test
    @DisplayName("ヘッダー一致 → ヘッダーは上書きされず、データがヘッダー行の次から書き込まれる")
    void writesAfterHeader() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "h1");
        setCell(sheet, 0, 1, "h2");

        new StringHeaderExcelTableWriter(
            "Sheet1", new String[]{"h1", "h2"}, 1, 1)
            .write(wb, List.of(List.of("d1", "d2")));

        assertThat(sheet.getRow(0).getCell(0).getStringCellValue()).isEqualTo("h1");
        assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("h2");
        assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("d1");
        assertThat(sheet.getRow(1).getCell(1).getStringCellValue()).isEqualTo("d2");
      }
    }

    @Test
    @DisplayName("ヘッダー不一致 → ExcelAppException（書き込みは行われない）")
    void headerMismatch() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "WRONG");
        setCell(sheet, 0, 1, "h2");

        StringHeaderExcelTableWriter writer = new StringHeaderExcelTableWriter(
            "Sheet1", new String[]{"h1", "h2"}, 1, 1);
        assertThatThrownBy(() -> writer.write(wb, List.of(List.of("d1", "d2"))))
            .isInstanceOf(ExcelAppException.class);
        assertThat(sheet.getRow(1)).isNull(); // data was not written
      }
    }
  }

  @Nested
  @DisplayName("複数行ヘッダーの書き込み")
  class MultiRowHeaderWrite {

    @Test
    @DisplayName("2行ヘッダーが書き込まれる")
    void writeTwoRowHeaders() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // pre-write template header matching what we'll validate against
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "個人情報");
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");

        new StringHeaderExcelTableWriter("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1)
            .write(wb, List.of(List.of("1", "Alice", "25")));

        // header rows are preserved
        assertThat(getCellValue(sheet, 0, 1)).isEqualTo("個人情報");
        assertThat(getCellValue(sheet, 1, 1)).isEqualTo("名前");
        // data row written after 2 header rows
        assertThat(getCellValue(sheet, 2, 0)).isEqualTo("1");
        assertThat(getCellValue(sheet, 2, 1)).isEqualTo("Alice");
        assertThat(getCellValue(sheet, 2, 2)).isEqualTo("25");
      }
    }

    @Test
    @DisplayName("writeHeaders で同行連続セルが横結合される")
    void horizontalMergeOnSameValues() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");

        new StringHeaderExcelTableWriter("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1)
            .writeHeaders(sheet);

        // "個人情報" spans cols 1-2 in row 0 → merged
        boolean merged = sheet.getMergedRegions().stream()
            .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 0
                && r.getFirstColumn() == 1 && r.getLastColumn() == 2);
        assertThat(merged).isTrue();
      }
    }

    @Test
    @DisplayName("writeHeaders で全行同値の列が縦結合される")
    void verticalMergeOnAllRowsSameValue() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");

        new StringHeaderExcelTableWriter("Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1)
            .writeHeaders(sheet);

        // "#" is same in both rows of col 0 → vertically merged (rows 0-1, col 0)
        boolean merged = sheet.getMergedRegions().stream()
            .anyMatch(r -> r.getFirstRow() == 0 && r.getLastRow() == 1
                && r.getFirstColumn() == 0 && r.getLastColumn() == 0);
        assertThat(merged).isTrue();
      }
    }
  }
}
