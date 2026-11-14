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
package jp.ecuacion.util.pdf.excel.report.util;

import static org.assertj.core.api.Assertions.assertThat;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Tests that bold text is emboldened (drawn with the text rendering mode "fill and stroke",
 * {@code 2 Tr}) when no bold font is available, in cells, headers / footers and shapes alike.
 */
@DisplayName("bold text without a bold font is emboldened")
class EmboldenedBoldTest {

  @TempDir
  Path tempDir;

  private Path generate(XSSFWorkbook wb, boolean withBoldFont) throws Exception {
    Path excel = tempDir.resolve("bold.xlsx");
    try (OutputStream out = Files.newOutputStream(excel)) {
      wb.write(out);
    }
    Path dir = Path.of(getClass().getResource("/fonts/NotoSansJP").toURI());
    PdfGenerateOptions.Builder builder =
        PdfGenerateOptions.builderForExplicitFont(dir.resolve("NotoSansJP-Regular.ttf"));
    if (withBoldFont) {
      builder.addBoldFontPath(dir.resolve("NotoSansJP-Bold.ttf"));
    }
    Path pdf = tempDir.resolve("bold.pdf");
    ExcelToPdfUtil.generate(excel, List.of("s"), pdf, builder.build());
    return pdf;
  }

  private static String content(Path pdf) throws Exception {
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      return new String(doc.getPage(0).getContents().readAllBytes(),
          StandardCharsets.ISO_8859_1);
    }
  }

  private static XSSFSheet sheetWithText(XSSFWorkbook wb, boolean boldCell) {
    XSSFSheet sheet = wb.createSheet("s");
    XSSFFont font = wb.createFont();
    font.setBold(boldCell);
    XSSFCellStyle style = wb.createCellStyle();
    style.setFont(font);
    var cell = sheet.createRow(0).createCell(0);
    cell.setCellStyle(style);
    cell.setCellValue("本文");
    wb.setPrintArea(0, 0, 0, 0, 0);
    return sheet;
  }

  @Test
  @DisplayName("cell: emboldened without a bold font, not with one")
  void cell() throws Exception {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      sheetWithText(wb, true);
      assertThat(content(generate(wb, false))).contains("\n2 Tr");
    }
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      sheetWithText(wb, true);
      assertThat(content(generate(wb, true))).doesNotContain("\n2 Tr");
    }
  }

  @Test
  @DisplayName("header and footer: &B text is emboldened without a bold font")
  void headerFooter() throws Exception {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      sheetWithText(wb, false).getHeader().setCenter("&B太字ヘッダー");
      assertThat(content(generate(wb, false))).contains("\n2 Tr");
    }
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      sheetWithText(wb, false).getHeader().setCenter("通常ヘッダー");
      assertThat(content(generate(wb, false))).doesNotContain("\n2 Tr");
    }
  }

  @Test
  @DisplayName("shape: bold text is emboldened without a bold font")
  void shape() throws Exception {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      XSSFSheet sheet = sheetWithText(wb, false);
      sheet.createRow(5);
      wb.setPrintArea(0, 0, 5, 0, 5);
      var drawing = sheet.createDrawingPatriarch();
      var shape = drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 1, 1, 4, 4));
      shape.setShapeType(org.apache.poi.ss.usermodel.ShapeTypes.RECT);
      var run = shape.addNewTextParagraph().addNewTextRun();
      run.setText("太字図形");
      run.setBold(true);
      assertThat(content(generate(wb, false))).contains("\n2 Tr");
    }
  }
}
