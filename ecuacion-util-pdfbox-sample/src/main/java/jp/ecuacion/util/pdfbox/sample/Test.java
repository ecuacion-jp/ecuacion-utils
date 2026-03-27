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
package jp.ecuacion.util.pdfbox.sample;

import java.io.FileInputStream;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdfbox.excel.exception.PdfGenerateException;
import jp.ecuacion.util.pdfbox.excel.util.ExcelToPdfUtil;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
  public static void main(String[] args) throws Exception {
    // Diagnose 2404関西 column widths
    try (var wb = (XSSFWorkbook) WorkbookFactory.create(
        new FileInputStream("/Users/yosuke.tanaka/Desktop/nencho.xlsx"))) {
      XSSFSheet sheet = wb.getSheet("2404関西");
      String printArea = wb.getPrintArea(wb.getSheetIndex(sheet));
      System.out.println("printArea=" + printArea);
      // parse lastCol from printArea
      String ref = printArea.substring(printArea.indexOf('!') + 1).replace("$", "");
      String[] parts = ref.split(":");
      int lastCol = cellRefToCol(parts[1]);
      int firstCol = cellRefToCol(parts[0]);
      System.out.println("firstCol=" + firstCol + " lastCol=" + lastCol);
      float total = 0;
      for (int c = firstCol; c <= lastCol; c++) {
        boolean hidden = sheet.isColumnHidden(c);
        float widthPx = sheet.getColumnWidthInPixels(c);
        float widthPt = widthPx * 72f / 96f;
        int rawWidth = sheet.getColumnWidth(c); // in 1/256 chars
        System.out.printf("col%d: hidden=%b rawWidth=%d widthPx=%.1f widthPt=%.2f%n",
            c, hidden, rawWidth, widthPx, widthPt);
        if (!hidden) total += widthPt;
      }
      System.out.printf("Total visible width=%.2fpt%n", total);
      double dcw = sheet.getCTWorksheet().isSetSheetFormatPr()
          ? sheet.getCTWorksheet().getSheetFormatPr().getDefaultColWidth() : -1;
      System.out.printf("defaultColWidth=%.6f defaultRowHeight=%.2f defaultColWidthPOI=%d%n",
          dcw, sheet.getDefaultRowHeightInPoints(), sheet.getDefaultColumnWidth());
    }

    ExcelToPdfUtil.generate(Path.of("/Users/yosuke.tanaka/Desktop/nencho.xlsx"), List.of("2404関西"),
        Path.of("/Users/yosuke.tanaka/Desktop/nencho.pdf"), null);
  }

  static int cellRefToCol(String ref) {
    int col = 0;
    for (char ch : ref.toCharArray()) {
      if (Character.isLetter(ch)) {
        col = col * 26 + (Character.toUpperCase(ch) - 'A' + 1);
      }
    }
    return col - 1;
  }
}
