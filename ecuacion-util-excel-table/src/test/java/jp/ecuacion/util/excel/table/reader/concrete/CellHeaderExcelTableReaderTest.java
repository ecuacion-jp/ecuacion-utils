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
import java.util.List;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

// 基底クラス共通の振る舞い（tableRowSize, 開始位置, isVerticalAndHorizontalOpposite,
// SheetNotExist 等）は StringFreeExcelTableReaderTest でカバー済み。
// ヘッダー検証の振る舞いは StringHeaderExcelTableReaderTest でカバー済み。
@DisplayName("CellHeaderExcelTableReader"
    + " ※基底クラス共通の振る舞いは StringFreeExcelTableReaderTest 参照"
    + "、ヘッダー検証は StringHeaderExcelTableReaderTest 参照")
public class CellHeaderExcelTableReaderTest {

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
  @DisplayName("Cell 型固有の振る舞い")
  class CellSpecific {

    @Test
    @DisplayName("ヘッダー行 + データ行 → ヘッダーは除外され Cell オブジェクトのリストで返る")
    void returnsCellObjectsWithoutHeader() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "header1");
        setCell(sheet, 0, 1, "header2");
        setCell(sheet, 1, 0, "data1-1");
        setCell(sheet, 1, 1, "data1-2");
        setCell(sheet, 2, 0, "data2-1");
        setCell(sheet, 2, 1, "data2-2");

        List<List<Cell>> result = new CellHeaderExcelTableReader(
            "Sheet1", new String[]{"header1", "header2"}).tableStartRowNumber(1).read(wb);

        assertThat(result).hasSize(2);
        assertThat(ExcelReadUtil.getStringFromCell(result.get(0).get(0))).isEqualTo("data1-1");
        assertThat(ExcelReadUtil.getStringFromCell(result.get(0).get(1))).isEqualTo("data1-2");
        assertThat(ExcelReadUtil.getStringFromCell(result.get(1).get(0))).isEqualTo("data2-1");
        assertThat(ExcelReadUtil.getStringFromCell(result.get(1).get(1))).isEqualTo("data2-2");
      }
    }
  }
}
