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
import java.util.List;
import jp.ecuacion.util.excel.table.reader.concrete.CellOneLineHeaderExcelTableReader;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

// 基底クラス共通の振る舞い（書き込み, 開始位置, isVerticalAndHorizontalOpposite, SheetNotExist 等）は
// StringFreeExcelTableWriterTest でカバー済み。
// ヘッダー検証の振る舞いは StringHeaderExcelTableWriterTest でカバー済み。
@DisplayName("CellOneLineHeaderExcelTableWriter"
    + " ※基底クラス共通の振る舞いは StringFreeExcelTableWriterTest 参照"
    + "、ヘッダー検証は StringHeaderExcelTableWriterTest 参照")
public class CellOneLineHeaderExcelTableWriterTest {

  private static void setCell(Sheet sheet, int poiRow, int poiCol, String value) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      row = sheet.createRow(poiRow);
    }
    row.createCell(poiCol).setCellValue(value);
  }

  @Nested
  @DisplayName("Cell 型固有の振る舞い")
  class CellSpecific {

    @Test
    @DisplayName("ヘッダー一致 → ヘッダーは上書きされず、Cell データがヘッダー行の次から書き込まれる")
    void writesCellsAfterHeader() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet src = wb.createSheet("source");
        setCell(src, 0, 0, "h1");
        setCell(src, 0, 1, "h2");
        setCell(src, 1, 0, "data1");
        setCell(src, 1, 1, "data2");

        Sheet dest = wb.createSheet("dest");
        setCell(dest, 0, 0, "h1");
        setCell(dest, 0, 1, "h2");

        List<List<Cell>> data = new CellOneLineHeaderExcelTableReader(
            "source", new String[]{"h1", "h2"}, 1, 1, null).read(wb);
        new CellOneLineHeaderExcelTableWriter(
            "dest", new String[]{"h1", "h2"}, 1, 1).write(wb, data);

        assertThat(dest.getRow(0).getCell(0).getStringCellValue()).isEqualTo("h1");
        assertThat(dest.getRow(0).getCell(1).getStringCellValue()).isEqualTo("h2");
        assertThat(ExcelReadUtil.getStringFromCell(dest.getRow(1).getCell(0)))
            .isEqualTo("data1");
        assertThat(ExcelReadUtil.getStringFromCell(dest.getRow(1).getCell(1)))
            .isEqualTo("data2");
      }
    }
  }
}
