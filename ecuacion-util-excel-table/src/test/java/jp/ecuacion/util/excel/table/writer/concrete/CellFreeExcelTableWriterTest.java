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
import jp.ecuacion.util.excel.table.reader.concrete.CellFreeExcelTableReader;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

// Common base-class behaviors (writing, start position, isVerticalAndHorizontalOpposite,
// SheetNotExist, etc.) are covered by StringFreeExcelTableWriterTest.
@DisplayName("CellFreeExcelTableWriter (base-class behaviors: see StringFreeExcelTableWriterTest)")
public class CellFreeExcelTableWriterTest {

  @Nested
  @DisplayName("Cell コピーの振る舞い")
  class CellCopy {

    @Test
    @DisplayName("Cell 値が正しくコピーされる")
    void copiesCellValue() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet src = wb.createSheet("source");
        src.createRow(0).createCell(0).setCellValue("hello");
        src.getRow(0).createCell(1).setCellValue("world");
        wb.createSheet("dest");

        List<List<Cell>> data =
            new CellFreeExcelTableReader("source").tableStartRowNumber(1).read(wb);
        new CellFreeExcelTableWriter("dest").tableStartRowNumber(1).write(wb, data);

        Sheet destSheet = wb.getSheet("dest");
        assertThat(ExcelReadUtil.getStringFromCell(destSheet.getRow(0).getCell(0)))
            .isEqualTo("hello");
        assertThat(ExcelReadUtil.getStringFromCell(destSheet.getRow(0).getCell(1)))
            .isEqualTo("world");
      }
    }

    @Test
    @DisplayName("copiesDataFormatOnly=false → ソースのフルスタイル（フォント等）がコピーされる")
    void copiesFullStyle() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet src = wb.createSheet("source");
        Cell srcCell = src.createRow(0).createCell(0);
        srcCell.setCellValue(1234.5);
        CellStyle srcStyle = wb.createCellStyle();
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        srcStyle.setFont(boldFont);
        srcStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        srcCell.setCellStyle(srcStyle);
        wb.createSheet("dest");

        List<List<Cell>> data =
            new CellFreeExcelTableReader("source").tableStartRowNumber(1).tableRowSize(1)
                .tableColumnSize(1).read(wb);
        new CellFreeExcelTableWriter("dest").tableStartRowNumber(1).write(wb, data);

        Cell destCell = wb.getSheet("dest").getRow(0).getCell(0);
        Font destFont = wb.getFontAt(destCell.getCellStyle().getFontIndex());
        assertThat(destFont.getBold()).isTrue();
      }
    }

    @Test
    @DisplayName("copiesDataFormatOnly=true → データフォーマットのみコピー、フォント等はコピーされない")
    void copiesDataFormatOnly() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet src = wb.createSheet("source");
        Cell srcCell = src.createRow(0).createCell(0);
        srcCell.setCellValue(1234.5);
        CellStyle srcStyle = wb.createCellStyle();
        Font boldFont = wb.createFont();
        boldFont.setBold(true);
        srcStyle.setFont(boldFont);
        srcStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        srcCell.setCellStyle(srcStyle);
        wb.createSheet("dest");

        List<List<Cell>> data =
            new CellFreeExcelTableReader("source").tableStartRowNumber(1).tableRowSize(1)
                .tableColumnSize(1).read(wb);
        new CellFreeExcelTableWriter("dest").tableStartRowNumber(1)
            .withCopiesDataFormatOnly(true).write(wb, data);

        Cell destCell = wb.getSheet("dest").getRow(0).getCell(0);
        Font destFont = wb.getFontAt(destCell.getCellStyle().getFontIndex());
        assertThat(destFont.getBold()).isFalse(); // bold not copied
        assertThat(destCell.getCellStyle().getDataFormatString()).isEqualTo("0.00");
      }
    }
  }
}
