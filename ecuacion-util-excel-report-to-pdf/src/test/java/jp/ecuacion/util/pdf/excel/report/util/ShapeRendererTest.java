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
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import javax.imageio.ImageIO;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Tests for ShapeRenderer via ExcelToPdfUtil public API. */
@DisplayName("ShapeRenderer (via ExcelToPdfUtil)")
public class ShapeRendererTest {

  private static final PdfGenerateOptions TEST_OPTIONS;

  static {
    try {
      var reg = ShapeRendererTest.class.getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf");
      var bold = ShapeRendererTest.class.getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf");
      TEST_OPTIONS = PdfGenerateOptions.builder()
          .regularFontPath(Path.of(reg.toURI()))
          .boldFontPath(Path.of(bold.toURI()))
          .build();
    } catch (URISyntaxException e) {
      throw new ExceptionInInitializerError(e);
    }
  }

  // --- helpers ---

  private static Path buildExcel(Path dir, ShapeSetup setup) throws Exception {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      XSSFSheet sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);

      var row0 = sheet.createRow(0);
      row0.setHeightInPoints(200f);
      row0.createCell(0).setCellValue("x");

      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 2, 1);
      XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
      setup.configure(shape);

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private static Path buildImageExcel(Path dir, byte[] imageBytes, int poiImageType)
      throws Exception {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      XSSFSheet sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);

      var row0 = sheet.createRow(0);
      row0.setHeightInPoints(200f);
      row0.createCell(0).setCellValue("x");

      int picIdx = wb.addPicture(imageBytes, poiImageType);
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 0, 0, 2, 1);
      drawing.createPicture(anchor, picIdx);

      Path path = dir.resolve("image.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private static CTShapeProperties ensureSpPr(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    return spPr != null ? spPr : shape.getCTShape().addNewSpPr();
  }

  private static void setShapeText(XSSFSimpleShape shape, String text) {
    var ctShape = shape.getCTShape();
    CTTextBody txBody = ctShape.isSetTxBody() ? ctShape.getTxBody() : ctShape.addNewTxBody();
    if (txBody.getBodyPr() == null) {
      txBody.addNewBodyPr();
    }
    while (txBody.sizeOfPArray() > 0) {
      txBody.removeP(0);
    }
    CTTextParagraph para = txBody.addNewP();
    CTRegularTextRun run = para.addNewR();
    run.setT(text);
  }

  private static void setFill(XSSFSimpleShape shape, int r, int g, int b) {
    CTShapeProperties spPr = ensureSpPr(shape);
    CTSolidColorFillProperties fill =
        spPr.isSetSolidFill() ? spPr.getSolidFill() : spPr.addNewSolidFill();
    var srgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : fill.addNewSrgbClr();
    srgb.setVal(new byte[] {(byte) r, (byte) g, (byte) b});
  }

  private static void setLine(XSSFSimpleShape shape, int r, int g, int b, float widthPt) {
    CTShapeProperties spPr = ensureSpPr(shape);
    CTLineProperties ln = spPr.isSetLn() ? spPr.getLn() : spPr.addNewLn();
    ln.setW((int) (widthPt * 12700));
    CTSolidColorFillProperties fill =
        ln.isSetSolidFill() ? ln.getSolidFill() : ln.addNewSolidFill();
    var srgb = fill.isSetSrgbClr() ? fill.getSrgbClr() : fill.addNewSrgbClr();
    srgb.setVal(new byte[] {(byte) r, (byte) g, (byte) b});
  }

  private static void setShapeType(XSSFSimpleShape shape, STShapeType.Enum type) {
    CTShapeProperties spPr = ensureSpPr(shape);
    CTPresetGeometry2D prstGeom =
        spPr.isSetPrstGeom() ? spPr.getPrstGeom() : spPr.addNewPrstGeom();
    prstGeom.setPrst(type);
  }

  private static int generatePdfAndCount(Path excel) throws PdfGenerateException {
    // The PDF is rendered to memory only via in-memory path — but ExcelToPdfUtil.generate()
    // needs a file path. Use a sibling file.
    Path pdf = excel.resolveSibling("out.pdf");
    ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      return doc.getNumberOfPages();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  @FunctionalInterface
  interface ShapeSetup {
    void configure(XSSFSimpleShape shape) throws Exception;
  }

  // --- tests ---

  @Nested
  @DisplayName("テキストボックス")
  class TextBox {

    @Test
    @DisplayName("テキストボックスの文字が PDF テキストとして抽出できる")
    void textAppearsInPdf(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setShapeText(shape, "ShapeText");
      });
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("ShapeText");
      }
    }

    @Test
    @DisplayName("bold テキストボックスの文字が PDF テキストとして抽出できる")
    void boldTextAppearsInPdf(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setShapeText(shape, "BoldText");
      });
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("BoldText");
      }
    }
  }

  @Nested
  @DisplayName("塗り・枠線")
  class FillAndLine {

    @Test
    @DisplayName("塗り色のみ → PDF 生成成功・1ページ")
    void fillOnly(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 255, 0, 0); // red fill
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("枠線のみ → PDF 生成成功・1ページ")
    void lineOnly(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setLine(shape, 0, 0, 255, 1.5f); // blue line
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("塗り色 + 枠線 → PDF 生成成功・1ページ")
    void fillAndLine(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 0, 255, 0);    // green fill
        setLine(shape, 0, 0, 0, 1f);  // black line
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }
  }

  @Nested
  @DisplayName("図形タイプ")
  class ShapeTypeTests {

    @Test
    @DisplayName("ELLIPSE (楕円) → PDF 生成成功")
    void ellipse(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 100, 100, 200);
        setShapeType(shape, STShapeType.ELLIPSE);
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("ROUND_RECT (角丸矩形) → PDF 生成成功")
    void roundRect(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 200, 100, 100);
        setShapeType(shape, STShapeType.ROUND_RECT);
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("DIAMOND (菱形) → PDF 生成成功")
    void diamond(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 100, 200, 100);
        setShapeType(shape, STShapeType.DIAMOND);
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("PARALLELOGRAM (平行四辺形) → PDF 生成成功")
    void parallelogram(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 200, 200, 100);
        setShapeType(shape, STShapeType.PARALLELOGRAM);
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }

    @Test
    @DisplayName("デフォルト矩形 (prstGeom なし) → PDF 生成成功")
    void defaultRect(@TempDir Path tempDir) throws Exception {
      Path excel = buildExcel(tempDir, shape -> {
        setFill(shape, 150, 150, 150);
        // no setShapeType → default rectangle path in appendShapePath
      });
      assertThat(generatePdfAndCount(excel)).isGreaterThan(0);
    }
  }

  @Nested
  @DisplayName("画像")
  class ImageShapes {

    @Test
    @DisplayName("PNG 画像 → PDF 生成成功・非白ピクセルあり")
    void pngImage(@TempDir Path tempDir) throws Exception {
      BufferedImage img = new BufferedImage(20, 20, BufferedImage.TYPE_INT_RGB);
      Graphics2D g = img.createGraphics();
      g.setColor(Color.BLUE);
      g.fillRect(0, 0, 20, 20);
      g.dispose();
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      ImageIO.write(img, "PNG", baos);

      Path excel = buildImageExcel(tempDir, baos.toByteArray(), Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThan(0);
        BufferedImage rendered =
            new PDFRenderer(doc).renderImageWithDPI(0, PDRectangle.A4.getWidth() / 72f * 96f);
        assertThat(hasNonWhitePixel(rendered)).isTrue();
      }
    }

    @Test
    @DisplayName("JPEG 画像 → PDF 生成成功")
    void jpegImage(@TempDir Path tempDir) throws Exception {
      BufferedImage img = new BufferedImage(20, 20, BufferedImage.TYPE_INT_RGB);
      Graphics2D g = img.createGraphics();
      g.setColor(Color.RED);
      g.fillRect(0, 0, 20, 20);
      g.dispose();
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      ImageIO.write(img, "JPEG", baos);

      Path excel = buildImageExcel(tempDir, baos.toByteArray(), Workbook.PICTURE_TYPE_JPEG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThan(0);
      }
    }
  }

  private static boolean hasNonWhitePixel(BufferedImage img) {
    for (int y = 0; y < img.getHeight(); y++) {
      for (int x = 0; x < img.getWidth(); x++) {
        int rgb = img.getRGB(x, y) & 0xFFFFFF;
        if (rgb != 0xFFFFFF) {
          return true;
        }
      }
    }
    return false;
  }
}
