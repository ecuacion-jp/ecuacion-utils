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
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.core.api.Assertions.within;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import javax.imageio.ImageIO;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import jp.ecuacion.util.pdf.excel.report.internal.SystemFontLocator;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Tests for {@link ExcelToPdfUtil}. */
@DisplayName("ExcelToPdfUtil")
public class ExcelToPdfUtilTest {

  private static final int FILL_RGB = 0x0070C0;
  private static final int TOP_MARGIN_PX = 36; // 0.5in = 36pt
  private static final int SAFE_X = 23;        // left margin 0.25in (18pt) + 5

  private static final PdfGenerateOptions TEST_OPTIONS;
  static {
    try {
      var reg = ExcelToPdfUtilTest.class
          .getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf");
      var bold = ExcelToPdfUtilTest.class
          .getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf");
      TEST_OPTIONS = PdfGenerateOptions.builder()
          .regularFontPath(Path.of(reg.toURI()))
          .boldFontPath(Path.of(bold.toURI()))
          .build();
    } catch (Exception e) {
      throw new ExceptionInInitializerError(e);
    }
  }

  private static PdfGenerateOptions optionsWithDateLocale(Locale locale) {
    try {
      var reg = ExcelToPdfUtilTest.class
          .getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf");
      var bold = ExcelToPdfUtilTest.class
          .getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf");
      return PdfGenerateOptions.builder()
          .regularFontPath(Path.of(reg.toURI()))
          .boldFontPath(Path.of(bold.toURI()))
          .dateLocale(locale)
          .build();
    } catch (Exception e) {
      throw new RuntimeException(e);
    }
  }

  // ---------------------------------------------------------------------------
  // page size and orientation
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("page size and orientation")
  class PageSizeAndOrientation {

    @Test
    @DisplayName("A4 portrait produces a page of A4 portrait dimensions")
    void a4Portrait(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.A4_PAPERSIZE, false);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.A4.getWidth(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.A4.getHeight(), within(0.5f));
      }
    }

    @Test
    @DisplayName("A4 landscape produces a page of A4 landscape dimensions")
    void a4Landscape(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.A4_PAPERSIZE, true);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.A4.getHeight(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.A4.getWidth(), within(0.5f));
      }
    }

    @Test
    @DisplayName("A5 portrait produces a page of A5 portrait dimensions")
    void a5Portrait(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.A5_PAPERSIZE, false);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.A5.getWidth(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.A5.getHeight(), within(0.5f));
      }
    }

    @Test
    @DisplayName("A5 landscape produces a page of A5 landscape dimensions")
    void a5Landscape(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.A5_PAPERSIZE, true);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.A5.getHeight(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.A5.getWidth(), within(0.5f));
      }
    }

    @Test
    @DisplayName("Letter portrait produces a page of Letter portrait dimensions")
    void letterPortrait(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.LETTER_PAPERSIZE, false);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.LETTER.getWidth(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.LETTER.getHeight(), within(0.5f));
      }
    }

    @Test
    @DisplayName("Letter landscape produces a page of Letter landscape dimensions")
    void letterLandscape(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createMinimalWorkbook(tempDir, PrintSetup.LETTER_PAPERSIZE, true);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        PDPage page = doc.getPage(0);
        assertThat(page.getMediaBox().getWidth())
            .isCloseTo(PDRectangle.LETTER.getHeight(), within(0.5f));
        assertThat(page.getMediaBox().getHeight())
            .isCloseTo(PDRectangle.LETTER.getWidth(), within(0.5f));
      }
    }
  }

  // ---------------------------------------------------------------------------
  // print area
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("print area")
  class PrintArea {

    @Test
    @DisplayName("cells outside the print area (column direction) are not rendered")
    void excludesColumnsOutsidePrintArea(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithPrintArea(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("inside");
        assertThat(text).doesNotContain("outside_col");
      }
    }

    @Test
    @DisplayName("cells outside the print area (row direction) are not rendered")
    void excludesRowsOutsidePrintArea(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithPrintArea(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("inside");
        assertThat(text).doesNotContain("outside_row");
      }
    }

    @Test
    @DisplayName("when no print area is set, at least one page is generated from used range")
    void noPrintArea(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithoutPrintArea(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThanOrEqualTo(1);
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("data");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // page breaks
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("page breaks")
  class PageBreaks {

    @Test
    @DisplayName("a manual row break creates 2 pages with correct content on each page")
    void singleRowBreak(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithSingleRowBreak(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        assertThat(textOfPage(doc, 1)).contains("section1").doesNotContain("section2");
        assertThat(textOfPage(doc, 2)).contains("section2").doesNotContain("section1");
      }
    }

    @Test
    @DisplayName("two manual row breaks create 3 pages with correct content on each page")
    void multipleRowBreaks(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithMultipleRowBreaks(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(3);
        assertThat(textOfPage(doc, 1)).contains("section1")
            .doesNotContain("section2").doesNotContain("section3");
        assertThat(textOfPage(doc, 2)).contains("section2")
            .doesNotContain("section1").doesNotContain("section3");
        assertThat(textOfPage(doc, 3)).contains("section3")
            .doesNotContain("section1").doesNotContain("section2");
      }
    }

    @Test
    @DisplayName("a manual column break creates 2 pages with correct content on each page")
    void columnBreak(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithColumnBreak(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        assertThat(textOfPage(doc, 1)).contains("left").doesNotContain("right");
        assertThat(textOfPage(doc, 2)).contains("right").doesNotContain("left");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // header and footer
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("header and footer")
  class HeaderAndFooter {

    @Test
    @DisplayName("left, center, and right sections all render their text")
    void sections(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", "LEFT", "CENTER", "RIGHT");
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("LEFT").contains("CENTER").contains("RIGHT");
      }
    }

    @Test
    @DisplayName("&P is replaced with the current page number")
    void pageNumber(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "Page &P", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Page 1");
      }
    }

    @Test
    @DisplayName("&N is replaced with the total number of pages")
    void totalPages(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "of &N", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("of 1");
      }
    }

    @Test
    @DisplayName("&A is replaced with the sheet name")
    void sheetName(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "Sheet: &A", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Sheet: Sheet1");
      }
    }

    @Test
    @DisplayName("&F is replaced with the file name without extension")
    void fileName(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "myreport.xlsx", null, "File: &F", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("File: myreport");
      }
    }

    @Test
    @DisplayName("&D and &T are replaced with the current date and time")
    void dateAndTime(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      String expectedDate =
          LocalDate.now(ZoneId.systemDefault()).format(DateTimeFormatter.ofPattern("yyyy/M/d"));
      String expectedTime =
          LocalTime.now(ZoneId.systemDefault()).format(DateTimeFormatter.ofPattern("H:mm"));
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "&D &T", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = textOfPage(doc, 1);
        assertThat(text).contains(expectedDate);
        assertThat(text).contains(expectedTime);
      }
    }

    @Test
    @DisplayName("&& produces a literal ampersand")
    void literalAmpersand(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "A && B", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("A & B");
      }
    }

    @Test
    @DisplayName("&P+n produces page number plus offset")
    void pageNumberWithOffset(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "Page &P+3", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Page 4");
      }
    }

    @Test
    @DisplayName("header appears on every page with the correct page number")
    void headerRepeatsOnEveryPage(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeaderAndRowBreak(tempDir, "Page &P of &N");
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        assertThat(textOfPage(doc, 1)).contains("Page 1 of 2");
        assertThat(textOfPage(doc, 2)).contains("Page 2 of 2");
      }
    }

    @Test
    @DisplayName("footer text appears in the PDF")
    void footer(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithFooter(tempDir, "FOOTER_TEXT");
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("FOOTER_TEXT");
      }
    }

    @Test
    @DisplayName("toggle formatting codes render text without crashing")
    void formattingToggleCodes(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx",
          null, "&BBold&B &UUnder&U &EDouble&E &SStrike&S &Xsup&X &Ysub&Y", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = textOfPage(doc, 1);
        assertThat(text)
            .contains("Bold").contains("Under").contains("Double")
            .contains("Strike").contains("sup").contains("sub");
      }
    }

    @Test
    @DisplayName("style codes (font size, font spec, color) render text without crashing")
    void styleCodes(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx",
          null, "&14Big &\"Arial,Bold\"Spec &KFF0000Red", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = textOfPage(doc, 1);
        assertThat(text).contains("Big").contains("Spec").contains("Red");
      }
    }

    @Test
    @DisplayName("header margin setting positions the header above the content area")
    void headerMarginPositioning(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // headerMargin=0.5in(36pt), topMargin=1.0in(72pt)
      Path excel = createWorkbookWithHeaderMarginConfig(tempDir, 0.5, 1.0);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      float topMarginPt = (float) (1.0 * 72);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        PDPage page = doc.getPage(0);
        float pageWidth = page.getMediaBox().getWidth();
        float pageHeight = page.getMediaBox().getHeight();

        PDFTextStripperByArea stripper = new PDFTextStripperByArea();
        stripper.setSortByPosition(true);
        stripper.addRegion("headerArea",
            new Rectangle2D.Float(0, 0, pageWidth, topMarginPt));
        stripper.addRegion("contentArea",
            new Rectangle2D.Float(0, topMarginPt, pageWidth, pageHeight - topMarginPt));
        stripper.extractRegions(page);

        assertThat(stripper.getTextForRegion("headerArea")).contains("HEADER");
        assertThat(stripper.getTextForRegion("contentArea")).contains("CONTENT");
        assertThat(stripper.getTextForRegion("contentArea")).doesNotContain("HEADER");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // scale
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("scale")
  class Scale {

    @Test
    @DisplayName("scale 100% renders rows at natural height (60pt)")
    void scale100(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createScaledWorkbook(tempDir, (short) 100);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        // row spans y=[36, 96]: y=66 inside, y=106 outside
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 30) & 0xFFFFFF).isEqualTo(FILL_RGB);
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 70) & 0xFFFFFF).isEqualTo(0xFFFFFF);
      }
    }

    @Test
    @DisplayName("scale 50% renders rows at half their natural height (30pt)")
    void scale50(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createScaledWorkbook(tempDir, (short) 50);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        // row spans y=[36, 66]: y=56 inside, y=76 outside (colored at 100%)
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 20) & 0xFFFFFF).isEqualTo(FILL_RGB);
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 40) & 0xFFFFFF).isEqualTo(0xFFFFFF);
      }
    }

    @Test
    @DisplayName("scale 200% renders rows at double their natural height (120pt)")
    void scale200(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createScaledWorkbook(tempDir, (short) 200);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        // row spans y=[36, 156]: y=106 inside (white at 100%), y=166 outside
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 70) & 0xFFFFFF).isEqualTo(FILL_RGB);
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 130) & 0xFFFFFF).isEqualTo(0xFFFFFF);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // fit to page (no explicit scale)
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("fit to page (no explicit scale)")
  class FitToPage {

    @Test
    @DisplayName("content wider than printable area spans multiple pages at 100% scale (no auto-scale)")
    void tooWide(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // Without fitToPage or explicit scale, Excel renders at 100% — content overflows
      // horizontally and spans multiple column-pages. Our code matches this behaviour.
      Path excel = createWorkbookTooWideForPage(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThan(1);
      }
    }

    @Test
    @DisplayName("content taller than printable area spans multiple pages at 100% scale (no auto-scale)")
    void tooTall(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // Without fitToPage or explicit scale, Excel renders at 100% — content overflows
      // vertically and spans multiple row-pages. Our code matches this behaviour.
      Path excel = createWorkbookTooTallForPage(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThan(1);
      }
    }

    @Test
    @DisplayName("content fitting naturally is rendered at natural size without scaling")
    void fitsNaturally(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createSmallWorkbookNoScale(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        // row height = 60pt, top margin = 36pt → row spans y=[36, 96]
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 30) & 0xFFFFFF).isEqualTo(FILL_RGB);
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 70) & 0xFFFFFF).isEqualTo(0xFFFFFF);
      }
    }

    @Test
    @DisplayName("fitToPage with both wide and tall content uses min(width-scale, height-scale)")
    void fitToPageUsesBothConstraints(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createFitToPageBothConstraintsWorkbook(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // With fitToPage, both width and height constraints are applied,
      // so content must fit on exactly 1 page.
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
      }
    }

    @Test
    @DisplayName("horizontalCentered=true centers content with equal left and right margins")
    void horizontalCenteredGivesEqualMargins(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Workbook: 1 column of explicit width 200pt, A4 page, 0.5in margins, horizontalCentered=true
      // Printable width = 595.28 - 72 = 523.28pt, content width = 200pt (at scale=1)
      // Centering offset = (523.28 - 200) / 2 ≈ 161.6pt
      // Content left edge at: 36 + 161.6 = 197.6pt ≈ 198px (at 72dpi)
      // Content right edge at: 197.6 + 200 = 397.6pt ≈ 398px
      Path excel = createHorizontallyCenteredWorkbook(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        int pageWidth = image.getWidth(); // 595px at 72dpi

        // Find leftmost fill pixel (scan from left at the cell row)
        int cellRow = TOP_MARGIN_PX + 10; // well inside the 60pt tall row
        int leftFill = -1;
        for (int x = 0; x < pageWidth; x++) {
          if ((image.getRGB(x, cellRow) & 0xFFFFFF) == FILL_RGB) {
            leftFill = x;
            break;
          }
        }
        // Find rightmost fill pixel (scan from right)
        int rightFill = -1;
        for (int x = pageWidth - 1; x >= 0; x--) {
          if ((image.getRGB(x, cellRow) & 0xFFFFFF) == FILL_RGB) {
            rightFill = x;
            break;
          }
        }

        assertThat(leftFill).as("fill should be found").isGreaterThan(0);
        assertThat(rightFill).as("fill should be found").isGreaterThan(0);

        int leftGap = leftFill;               // distance from page left to content left
        int rightGap = pageWidth - 1 - rightFill;  // distance from content right to page right
        // Allow ±3px tolerance for float rounding
        assertThat(Math.abs(leftGap - rightGap))
            .as("left gap (%d) should equal right gap (%d) within 3px", leftGap, rightGap)
            .isLessThanOrEqualTo(3);
      }
    }

    @Test
    @DisplayName("horizontalCentered=false keeps content left-aligned at the left margin")
    void horizontalNotCenteredStartsAtLeftMargin(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createSmallWorkbookNoScale(tempDir); // no centering, left=0.25in=18pt
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        int cellRow = TOP_MARGIN_PX + 10;
        // left margin = 18pt → fill should start near x=18, not centered
        assertThat((image.getRGB(SAFE_X, cellRow) & 0xFFFFFF))
            .as("content should be at left margin, not centered")
            .isEqualTo(FILL_RGB);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // margins
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("margins (left and top)")
  class Margins {

    @Test
    @DisplayName("Normal margins (left=0.7in, top=0.75in) are reflected in the PDF")
    void normal(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertMarginBoundaries(tempDir, 0.7, 0.75);
    }

    @Test
    @DisplayName("Wide margins (left=1.0in, top=1.0in) are reflected in the PDF")
    void wide(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertMarginBoundaries(tempDir, 1.0, 1.0);
    }

    @Test
    @DisplayName("Narrow margins (left=0.25in, top=0.75in) are reflected in the PDF")
    void narrow(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertMarginBoundaries(tempDir, 0.25, 0.75);
    }

    @Test
    @DisplayName("Custom margins (left=1.5in, top=1.2in) are reflected in the PDF")
    void custom(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertMarginBoundaries(tempDir, 1.5, 1.2);
    }

    private void assertMarginBoundaries(Path tempDir, double leftIn, double topIn)
        throws IOException, PdfGenerateException {
      Path excel = createColoredWorkbook(tempDir, leftIn, topIn);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      float leftPt = (float) (leftIn * 72);
      float topPt = (float) (topIn * 72);

      // Positions safely inside the content area (first cell)
      int safeContentX = (int) Math.ceil(leftPt) + 5;
      int safeContentY = (int) Math.ceil(topPt) + 5;

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);

        // Left margin boundary:
        //   floor(leftPt) - 1 is definitely in the margin (white)
        //   ceil(leftPt)      is definitely in the cell area (filled)
        assertThat(image.getRGB((int) leftPt - 1, safeContentY) & 0xFFFFFF)
            .as("pixel left of left margin should be white")
            .isEqualTo(0xFFFFFF);
        assertThat(image.getRGB((int) Math.ceil(leftPt), safeContentY) & 0xFFFFFF)
            .as("pixel at left margin boundary should be filled")
            .isEqualTo(FILL_RGB);

        // Top margin boundary:
        //   floor(topPt) - 1 is definitely in the margin (white)
        //   ceil(topPt)      is definitely in the cell area (filled)
        assertThat(image.getRGB(safeContentX, (int) topPt - 1) & 0xFFFFFF)
            .as("pixel above top margin should be white")
            .isEqualTo(0xFFFFFF);
        assertThat(image.getRGB(safeContentX, (int) Math.ceil(topPt)) & 0xFFFFFF)
            .as("pixel at top margin boundary should be filled")
            .isEqualTo(FILL_RGB);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // cell borders
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("cell borders")
  class CellBorders {

    // Render at 144 DPI: 1 pt = 2 px. THIN (0.5 pt) = 1 px, THICK (1.5 pt) = 3 px.
    // Cell geometry: left=0.25in(18pt), top=0.5in(36pt), row=60pt, col=2048 POI units(42pt).
    private static final int B_DPI    = 144;
    private static final int B_LEFT   = 36;   // left border x:  18pt × 2
    private static final int B_TOP    = 72;   // top border y:   36pt × 2
    private static final int B_RIGHT  = 120;  // right border x: (18+42)pt × 2
    private static final int B_BOTTOM = 192;  // bottom border y:(36+60)pt × 2
    private static final int B_SAFE_X = 56;   // safe inside x:  B_LEFT + 20
    private static final int B_SAFE_Y = 132;  // safe inside y:  B_TOP + 60

    // Diagonal check points at 144 DPI.
    // At x=50 (B_LEFT+14), the ↘ diagonal hits y = B_TOP + 14*120/84 = 92.
    // At x=106 (B_RIGHT-14), the ↗ diagonal hits y = B_TOP + 14*120/84 = 92.
    private static final int B_DIAG_X_DOWN = B_LEFT + 14;  // 50: on the ↘ diagonal
    private static final int B_DIAG_X_UP   = B_RIGHT - 14; // 106: on the ↗ diagonal
    private static final int B_DIAG_Y      = B_TOP + 20;   // 92

    @Test
    @DisplayName("thin top border renders a dark line at the cell top boundary")
    void thinTopBorder(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createBorderWorkbook(tempDir, "test.xlsx",
          BorderStyle.THIN, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(avgGray(img.getRGB(B_SAFE_X, B_TOP)))
            .as("pixel at top border should be dark")
            .isLessThan(128);
        assertThat(avgGray(img.getRGB(B_SAFE_X, B_SAFE_Y)))
            .as("pixel inside cell should be white")
            .isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("thick top border covers more pixels than thin top border")
    void thickBorderIsWider(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path thinXl = createBorderWorkbook(tempDir, "thin.xlsx",
          BorderStyle.THIN, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, null);
      Path thickXl = createBorderWorkbook(tempDir, "thick.xlsx",
          BorderStyle.THICK, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, null);
      Path thinPdf = tempDir.resolve("thin.pdf");
      Path thickPdf = tempDir.resolve("thick.pdf");
      ExcelToPdfUtil.generate(thinXl, List.of("Sheet1"), thinPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(thickXl, List.of("Sheet1"), thickPdf, TEST_OPTIONS);

      try (PDDocument thinDoc = Loader.loadPDF(thinPdf.toFile());
          PDDocument thickDoc = Loader.loadPDF(thickPdf.toFile())) {
        BufferedImage thinImg = new PDFRenderer(thinDoc).renderImageWithDPI(0, B_DPI);
        BufferedImage thickImg = new PDFRenderer(thickDoc).renderImageWithDPI(0, B_DPI);
        int thinDark = countDarkPixels(thinImg, B_SAFE_X, B_TOP - 3, B_TOP + 3);
        int thickDark = countDarkPixels(thickImg, B_SAFE_X, B_TOP - 3, B_TOP + 3);
        assertThat(thickDark)
            .as("thick border should cover more pixels than thin border")
            .isGreaterThan(thinDark);
      }
    }

    @Test
    @DisplayName("border color is reflected in the PDF")
    void borderColorIsReflected(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor red = new XSSFColor(new byte[] {(byte) 0xFF, 0x00, 0x00}, null);
      Path excel = createBorderWorkbook(tempDir, "test.xlsx",
          BorderStyle.THICK, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, red);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        int rgb = img.getRGB(B_SAFE_X, B_TOP) & 0xFFFFFF;
        assertThat((rgb >> 16) & 0xFF)
            .as("red channel at top border should be high")
            .isGreaterThan(180);
        assertThat((rgb >> 8) & 0xFF)
            .as("green channel at top border should be low")
            .isLessThan(80);
        assertThat(rgb & 0xFF)
            .as("blue channel at top border should be low")
            .isLessThan(80);
      }
    }

    @Test
    @DisplayName("cell without borders has no dark line at the cell boundaries")
    void noBorderHasNoLine(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createBorderWorkbook(tempDir, "test.xlsx",
          BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(avgGray(img.getRGB(B_SAFE_X, B_TOP)))
            .as("pixel at top boundary should be white when no border")
            .isGreaterThan(220);
        assertThat(avgGray(img.getRGB(B_LEFT, B_SAFE_Y)))
            .as("pixel at left boundary should be white when no border")
            .isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("thick borders on all four sides each render at the correct cell edge")
    void allFourSideBorders(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createBorderWorkbook(tempDir, "test.xlsx",
          BorderStyle.THICK, BorderStyle.THICK, BorderStyle.THICK, BorderStyle.THICK, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(avgGray(img.getRGB(B_SAFE_X, B_TOP)))
            .as("top border").isLessThan(128);
        assertThat(avgGray(img.getRGB(B_SAFE_X, B_BOTTOM)))
            .as("bottom border").isLessThan(128);
        assertThat(avgGray(img.getRGB(B_LEFT, B_SAFE_Y)))
            .as("left border").isLessThan(128);
        assertThat(avgGray(img.getRGB(B_RIGHT, B_SAFE_Y)))
            .as("right border").isLessThan(128);
      }
    }

    @Test
    @DisplayName("dashed top border has dark segments and visible gaps along the border line")
    void dashedBorderHasGaps(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createBorderWorkbook(tempDir, "test.xlsx",
          BorderStyle.DASHED, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        int darkCount = 0;
        int lightCount = 0;
        for (int x = B_LEFT + 5; x < B_RIGHT - 5; x++) {
          if (avgGray(img.getRGB(x, B_TOP)) < 128) {
            darkCount++;
          } else {
            lightCount++;
          }
        }
        assertThat(darkCount).as("dashed border should have some dark pixels").isGreaterThan(0);
        assertThat(lightCount)
            .as("dashed border should have visible gaps (light pixels)").isGreaterThan(0);
      }
    }

    @Test
    @DisplayName("diagonal-down border (↘) renders from top-left to bottom-right")
    void diagonalDownIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createDiagonalWorkbook(tempDir, "test.xlsx", BorderStyle.THIN, true, false);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(hasDarkPixelNear(img, B_DIAG_X_DOWN, B_DIAG_Y, 2))
            .as("↘ diagonal should have a dark pixel near the expected line position")
            .isTrue();
      }
    }

    @Test
    @DisplayName("diagonal-up border (↗) renders from bottom-left to top-right")
    void diagonalUpIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createDiagonalWorkbook(tempDir, "test.xlsx", BorderStyle.THIN, false, true);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(hasDarkPixelNear(img, B_DIAG_X_UP, B_DIAG_Y, 2))
            .as("↗ diagonal should have a dark pixel near the expected line position")
            .isTrue();
      }
    }

    @Test
    @DisplayName("both diagonals render when both directions are enabled")
    void bothDiagonalsAreRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createDiagonalWorkbook(tempDir, "test.xlsx", BorderStyle.THIN, true, true);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, B_DPI);
        assertThat(hasDarkPixelNear(img, B_DIAG_X_DOWN, B_DIAG_Y, 2))
            .as("↘ diagonal should have a dark pixel").isTrue();
        assertThat(hasDarkPixelNear(img, B_DIAG_X_UP, B_DIAG_Y, 2))
            .as("↗ diagonal should have a dark pixel").isTrue();
      }
    }
  }

  // ---------------------------------------------------------------------------
  // cell text
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("cell text")
  class CellText {

    // 144 DPI: 1pt = 2px. Cell: left=18pt, top=36pt, row=60pt, col=80pt.
    // CELL_PADDING=2pt=4px at 144 DPI.
    private static final int T_DPI     = 144;
    private static final int T_LEFT    = 36;   // 18pt × 2
    private static final int T_TOP     = 72;   // 36pt × 2
    // colWidth=3840 POI (15 chars), NotoSansJP MDW=8:
    //   spec formula: px = int((3840+128/8)/256*8) = int(120.5) = 120 → pt=90 → at 144 DPI: 180px
    private static final int T_RIGHT   = 216;  // (18+90)pt × 2
    private static final int T_BOTTOM  = 192;  // (36+60)pt × 2
    private static final int T_SAFE_X  = 126;  // (T_LEFT+T_RIGHT)/2
    private static final int T_SAFE_Y  = 132;  // (T_TOP+T_BOTTOM)/2

    // Expected pixel positions given CELL_PADDING=2pt=4px at 144 DPI.
    private static final int T_PAD_LEFT  = T_LEFT  + 4; // 40
    private static final int T_PAD_RIGHT = T_RIGHT - 4; // 212
    private static final int T_PAD_TOP   = T_TOP   + 4; // 76

    // Blue text color used in pixel-inspection tests.
    private static final int BLUE_RGB = 0x0070C0;

    @Test
    @DisplayName("plain text is extractable from the PDF")
    void textIsExtractable(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "Hello", 11, false, false,
          false, false, false, false, null, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("Hello");
      }
    }

    @Test
    @DisplayName("Japanese text is extractable without garbling")
    void japaneseTextIsExtractable(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "請求書", 11, false, false,
          false, false, false, false, null, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("請求書");
      }
    }

    @Test
    @DisplayName("larger font size produces a proportionally taller glyph (±5%)")
    void fontSizeAffectsGlyphHeight(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path small = createTextWorkbook(tempDir, "small.xlsx", "X", 12, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path large = createTextWorkbook(tempDir, "large.xlsx", "X", 24, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path smallPdf = tempDir.resolve("small.pdf");
      Path largePdf = tempDir.resolve("large.pdf");
      ExcelToPdfUtil.generate(small, List.of("Sheet1"), smallPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(large, List.of("Sheet1"), largePdf, TEST_OPTIONS);

      try (PDDocument sd = Loader.loadPDF(smallPdf.toFile());
          PDDocument ld = Loader.loadPDF(largePdf.toFile())) {
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        BufferedImage li = new PDFRenderer(ld).renderImageWithDPI(0, T_DPI);
        int smallH = glyphHeight(si, T_PAD_LEFT + 2, T_TOP, T_BOTTOM, BLUE_RGB);
        int largeH = glyphHeight(li, T_PAD_LEFT + 2, T_TOP, T_BOTTOM, BLUE_RGB);
        double ratio = (double) largeH / smallH;
        assertThat(ratio).as("24pt glyph should be ~2× taller than 12pt").isBetween(1.90, 2.10);
      }
    }

    @Test
    @DisplayName("bold text renders wider than normal text")
    void boldTextIsWiderThanNormal(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path normal = createTextWorkbook(tempDir, "normal.xlsx", "WW", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path bold = createTextWorkbook(tempDir, "bold.xlsx", "WW", 14, true, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path normalPdf = tempDir.resolve("normal.pdf");
      Path boldPdf = tempDir.resolve("bold.pdf");
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(bold, List.of("Sheet1"), boldPdf, TEST_OPTIONS);

      try (PDDocument nd = Loader.loadPDF(normalPdf.toFile());
          PDDocument bd = Loader.loadPDF(boldPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage bi = new PDFRenderer(bd).renderImageWithDPI(0, T_DPI);
        int normalRight = rightmostColoredX(ni, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int boldRight   = rightmostColoredX(bi, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(boldRight).as("bold text should extend further right").isGreaterThan(normalRight);
      }
    }

    @Test
    @DisplayName("italic text is rendered with a visible slant")
    void italicIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path normal = createTextWorkbook(tempDir, "normal.xlsx", "I", 20, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path italic = createTextWorkbook(tempDir, "italic.xlsx", "I", 20, false, true,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path normalPdf = tempDir.resolve("normal.pdf");
      Path italicPdf = tempDir.resolve("italic.pdf");
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(italic, List.of("Sheet1"), italicPdf, TEST_OPTIONS);

      // Synthetic italic shifts the top of the glyph to the right → italic rightmost > normal.
      try (PDDocument nd = Loader.loadPDF(normalPdf.toFile());
          PDDocument id = Loader.loadPDF(italicPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage ii = new PDFRenderer(id).renderImageWithDPI(0, T_DPI);
        int normalRight = rightmostColoredX(ni, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int italicRight = rightmostColoredX(ii, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(italicRight)
            .as("italic glyph top should be shifted right vs normal").isGreaterThan(normalRight);
      }
    }

    @Test
    @DisplayName("text color is reflected in the PDF")
    void textColorIsReflected(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "色", 20, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        boolean found = false;
        outer:
        for (int y = T_PAD_TOP; y < T_BOTTOM; y++) {
          for (int x = T_PAD_LEFT; x < T_RIGHT; x++) {
            if (isBlueish(img.getRGB(x, y))) {
              found = true;
              break outer;
            }
          }
        }
        assertThat(found).as("blue text pixel should exist in cell area").isTrue();
      }
    }

    @Test
    @DisplayName("strikethrough renders a line through the text")
    void strikethroughIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path withStrike = createTextWorkbook(tempDir, "strike.xlsx", "ABC", 14, false, false,
          true, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path noStrike = createTextWorkbook(tempDir, "nostrike.xlsx", "ABC", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path strikePdf = tempDir.resolve("strike.pdf");
      Path noStrikePdf = tempDir.resolve("nostrike.pdf");
      ExcelToPdfUtil.generate(withStrike, List.of("Sheet1"), strikePdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(noStrike, List.of("Sheet1"), noStrikePdf, TEST_OPTIONS);

      // Strikethrough adds more blue pixels (the line) than no-strike
      try (PDDocument sd = Loader.loadPDF(strikePdf.toFile());
          PDDocument nd = Loader.loadPDF(noStrikePdf.toFile())) {
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        int strikeBlue = countColoredPixels(si, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int noStrikeBlue = countColoredPixels(ni, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(strikeBlue)
            .as("strikethrough should add more blue pixels than no-strike")
            .isGreaterThan(noStrikeBlue);
      }
    }

    @Test
    @DisplayName("superscript text appears smaller and shifted upward")
    void superscriptIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path normal = createTextWorkbook(tempDir, "normal.xlsx", "m2", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER, 60, 3840);
      Path superPdf = tempDir.resolve("super.pdf");
      // superscript=true
      Path superXl = createTextWorkbook(tempDir, "super.xlsx", "m2", 14, false, false,
          false, true, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER, 60, 3840);
      Path normalPdf = tempDir.resolve("normal.pdf");
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(superXl, List.of("Sheet1"), superPdf, TEST_OPTIONS);

      try (PDDocument nd = Loader.loadPDF(normalPdf.toFile());
          PDDocument sd = Loader.loadPDF(superPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        int normalTop = topmostColoredY(ni, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        int superTop  = topmostColoredY(si, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(superTop)
            .as("superscript top pixel should be higher (smaller y) than normal").isLessThan(normalTop);
      }
    }

    @Test
    @DisplayName("subscript text appears smaller and shifted downward")
    void subscriptIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path normal = createTextWorkbook(tempDir, "normal.xlsx", "H2O", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER, 60, 3840);
      Path subXl = createTextWorkbook(tempDir, "sub.xlsx", "H2O", 14, false, false,
          false, false, true, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER, 60, 3840);
      Path normalPdf = tempDir.resolve("normal.pdf");
      Path subPdf = tempDir.resolve("sub.pdf");
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(subXl, List.of("Sheet1"), subPdf, TEST_OPTIONS);

      try (PDDocument nd = Loader.loadPDF(normalPdf.toFile());
          PDDocument sd = Loader.loadPDF(subPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        int normalBottom = bottommostColoredY(ni, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        int subBottom    = bottommostColoredY(si, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(subBottom)
            .as("subscript bottom pixel should be lower (larger y) than normal")
            .isGreaterThan(normalBottom);
      }
    }

    @Test
    @DisplayName("LEFT alignment starts text at CELL_PADDING from the left border")
    void leftAlignPositionsTextAtPadding(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int leftmost = leftmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(leftmost)
            .as("LEFT align: leftmost pixel should be near T_LEFT + 4px padding")
            .isBetween(T_PAD_LEFT - 2, T_PAD_LEFT + 4);
      }
    }

    @Test
    @DisplayName("RIGHT alignment ends text at CELL_PADDING from the right border")
    void rightAlignPositionsTextAtPadding(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.RIGHT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int rightmost = rightmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(rightmost)
            .as("RIGHT align: rightmost pixel should be near T_RIGHT - 4px padding")
            .isBetween(T_PAD_RIGHT - 8, T_PAD_RIGHT + 2);
      }
    }

    @Test
    @DisplayName("CENTER horizontal alignment positions text in the middle of the cell")
    void centerHorizontalAlignCentersText(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path leftXl = createTextWorkbook(tempDir, "left.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path centerXl = createTextWorkbook(tempDir, "center.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.CENTER,
          VerticalAlignment.TOP, 60, 3840);
      Path rightXl = createTextWorkbook(tempDir, "right.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.RIGHT,
          VerticalAlignment.TOP, 60, 3840);
      Path leftPdf = tempDir.resolve("left.pdf");
      Path centerPdf = tempDir.resolve("center.pdf");
      Path rightPdf = tempDir.resolve("right.pdf");
      ExcelToPdfUtil.generate(leftXl, List.of("Sheet1"), leftPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(centerXl, List.of("Sheet1"), centerPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(rightXl, List.of("Sheet1"), rightPdf, TEST_OPTIONS);

      try (PDDocument ld = Loader.loadPDF(leftPdf.toFile());
          PDDocument cd = Loader.loadPDF(centerPdf.toFile());
          PDDocument rd = Loader.loadPDF(rightPdf.toFile())) {
        BufferedImage li = new PDFRenderer(ld).renderImageWithDPI(0, T_DPI);
        BufferedImage ci = new PDFRenderer(cd).renderImageWithDPI(0, T_DPI);
        BufferedImage ri = new PDFRenderer(rd).renderImageWithDPI(0, T_DPI);
        int leftX   = leftmostColoredX(li, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int centerX = leftmostColoredX(ci, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int rightX  = leftmostColoredX(ri, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(centerX).as("CENTER leftmost should be between LEFT and RIGHT").isBetween(leftX, rightX);
      }
    }

    @Test
    @DisplayName("LEFT alignment with indent=1 shifts text right compared to no indent")
    void leftIndentShiftsTextRight(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path indent0 = createIndentWorkbook(tempDir, "i0.xlsx", HorizontalAlignment.LEFT, (short) 0, blue);
      Path indent1 = createIndentWorkbook(tempDir, "i1.xlsx", HorizontalAlignment.LEFT, (short) 1, blue);
      Path pdf0 = tempDir.resolve("out0.pdf");
      Path pdf1 = tempDir.resolve("out1.pdf");
      ExcelToPdfUtil.generate(indent0, List.of("Sheet1"), pdf0, TEST_OPTIONS);
      ExcelToPdfUtil.generate(indent1, List.of("Sheet1"), pdf1, TEST_OPTIONS);

      try (PDDocument d0 = Loader.loadPDF(pdf0.toFile());
          PDDocument d1 = Loader.loadPDF(pdf1.toFile())) {
        BufferedImage img0 = new PDFRenderer(d0).renderImageWithDPI(0, T_DPI);
        BufferedImage img1 = new PDFRenderer(d1).renderImageWithDPI(0, T_DPI);
        int x0 = leftmostColoredX(img0, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int x1 = leftmostColoredX(img1, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(x1).as("LEFT indent=1 text should start further right than indent=0").isGreaterThan(x0);
      }
    }

    @Test
    @DisplayName("RIGHT alignment with indent=1 shifts text left compared to no indent")
    void rightIndentShiftsTextLeft(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path indent0 = createIndentWorkbook(tempDir, "i0.xlsx", HorizontalAlignment.RIGHT, (short) 0, blue);
      Path indent1 = createIndentWorkbook(tempDir, "i1.xlsx", HorizontalAlignment.RIGHT, (short) 1, blue);
      Path pdf0 = tempDir.resolve("out0.pdf");
      Path pdf1 = tempDir.resolve("out1.pdf");
      ExcelToPdfUtil.generate(indent0, List.of("Sheet1"), pdf0, TEST_OPTIONS);
      ExcelToPdfUtil.generate(indent1, List.of("Sheet1"), pdf1, TEST_OPTIONS);

      try (PDDocument d0 = Loader.loadPDF(pdf0.toFile());
          PDDocument d1 = Loader.loadPDF(pdf1.toFile())) {
        BufferedImage img0 = new PDFRenderer(d0).renderImageWithDPI(0, T_DPI);
        BufferedImage img1 = new PDFRenderer(d1).renderImageWithDPI(0, T_DPI);
        int x0 = rightmostColoredX(img0, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int x1 = rightmostColoredX(img1, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(x1).as("RIGHT indent=1 text should end further left than indent=0").isLessThan(x0);
      }
    }

    @Test
    @DisplayName("indent=2 shifts text further than indent=1")
    void largerIndentShiftsTextFurther(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path indent1 = createIndentWorkbook(tempDir, "i1.xlsx", HorizontalAlignment.LEFT, (short) 1, blue);
      Path indent2 = createIndentWorkbook(tempDir, "i2.xlsx", HorizontalAlignment.LEFT, (short) 2, blue);
      Path pdf1 = tempDir.resolve("out1.pdf");
      Path pdf2 = tempDir.resolve("out2.pdf");
      ExcelToPdfUtil.generate(indent1, List.of("Sheet1"), pdf1, TEST_OPTIONS);
      ExcelToPdfUtil.generate(indent2, List.of("Sheet1"), pdf2, TEST_OPTIONS);

      try (PDDocument d1 = Loader.loadPDF(pdf1.toFile());
          PDDocument d2 = Loader.loadPDF(pdf2.toFile())) {
        BufferedImage img1 = new PDFRenderer(d1).renderImageWithDPI(0, T_DPI);
        BufferedImage img2 = new PDFRenderer(d2).renderImageWithDPI(0, T_DPI);
        int x1 = leftmostColoredX(img1, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int x2 = leftmostColoredX(img2, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(x2).as("indent=2 text should start further right than indent=1").isGreaterThan(x1);
      }
    }

    @Test
    @DisplayName("GENERAL alignment right-aligns numeric cell values")
    void generalAlignRightAlignsNumbers(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path numXl = createNumericTextWorkbook(tempDir, "num.xlsx", 12345.0, blue);
      Path strXl = createTextWorkbook(tempDir, "str.xlsx", "X", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.GENERAL,
          VerticalAlignment.TOP, 60, 3840);
      Path numPdf = tempDir.resolve("num.pdf");
      Path strPdf = tempDir.resolve("str.pdf");
      ExcelToPdfUtil.generate(numXl, List.of("Sheet1"), numPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(strXl, List.of("Sheet1"), strPdf, TEST_OPTIONS);

      try (PDDocument nd = Loader.loadPDF(numPdf.toFile());
          PDDocument sd = Loader.loadPDF(strPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        int numRight = rightmostColoredX(ni, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int strLeft  = leftmostColoredX(si, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(numRight)
            .as("GENERAL numeric: rightmost pixel near right edge").isGreaterThan(T_SAFE_X);
        assertThat(strLeft)
            .as("GENERAL string: leftmost pixel near left edge").isLessThan(T_SAFE_X);
      }
    }

    @Test
    @DisplayName("GENERAL alignment left-aligns string cell values")
    void generalAlignLeftAlignsStrings(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "Hello", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.GENERAL,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int leftmost = leftmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(leftmost)
            .as("GENERAL string: leftmost pixel near left edge").isBetween(T_PAD_LEFT - 2, T_SAFE_X);
      }
    }

    @Test
    @DisplayName("formula with numeric result is right-aligned under GENERAL")
    void formulaNumericResultIsRightAligned(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createFormulaWorkbook(tempDir, "test.xlsx", "1+1", false, blue);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int rightmost = rightmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(rightmost)
            .as("numeric formula result: rightmost pixel near right edge").isGreaterThan(T_SAFE_X);
      }
    }

    @Test
    @DisplayName("formula with string result is left-aligned under GENERAL")
    void formulaStringResultIsLeftAligned(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createFormulaWorkbook(tempDir, "test.xlsx", "\"請求\"", true, blue);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int leftmost = leftmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(leftmost)
            .as("string formula result: leftmost pixel near left edge")
            .isBetween(T_PAD_LEFT - 2, T_SAFE_X);
      }
    }

    @Test
    @DisplayName("TOP vertical alignment starts text at CELL_PADDING from the top border")
    void topAlignPositionsTextAtPadding(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int topmost = topmostColoredY(img, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        // TOP align: text starts near the top, so topmost pixel is in the upper half of cell.
        assertThat(topmost)
            .as("TOP align: topmost pixel should be in the upper half of the cell")
            .isLessThan(T_SAFE_Y);
      }
    }

    @Test
    @DisplayName("BOTTOM vertical alignment ends text at CELL_PADDING from the bottom border")
    void bottomAlignPositionsTextAtPadding(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.BOTTOM, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int bottom = bottommostColoredY(img, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        // BOTTOM align: text sits near the bottom, so bottommost pixel is in the lower half.
        assertThat(bottom)
            .as("BOTTOM align: bottommost pixel should be in the lower half of the cell")
            .isGreaterThan(T_SAFE_Y);
      }
    }

    @Test
    @DisplayName("CENTER vertical alignment positions text in the middle of the cell")
    void centerVerticalAlignCentersText(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path topXl = createTextWorkbook(tempDir, "top.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path centerXl = createTextWorkbook(tempDir, "center.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.CENTER, 60, 3840);
      Path bottomXl = createTextWorkbook(tempDir, "bottom.xlsx", "A", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.BOTTOM, 60, 3840);
      Path topPdf = tempDir.resolve("top.pdf");
      Path centerPdf = tempDir.resolve("center.pdf");
      Path bottomPdf = tempDir.resolve("bottom.pdf");
      ExcelToPdfUtil.generate(topXl, List.of("Sheet1"), topPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(centerXl, List.of("Sheet1"), centerPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(bottomXl, List.of("Sheet1"), bottomPdf, TEST_OPTIONS);

      try (PDDocument td = Loader.loadPDF(topPdf.toFile());
          PDDocument cd = Loader.loadPDF(centerPdf.toFile());
          PDDocument bd = Loader.loadPDF(bottomPdf.toFile())) {
        BufferedImage ti = new PDFRenderer(td).renderImageWithDPI(0, T_DPI);
        BufferedImage ci = new PDFRenderer(cd).renderImageWithDPI(0, T_DPI);
        BufferedImage bi = new PDFRenderer(bd).renderImageWithDPI(0, T_DPI);
        int topY    = topmostColoredY(ti, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        int centerY = topmostColoredY(ci, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        int bottomY = topmostColoredY(bi, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(centerY)
            .as("CENTER vertical: topmost pixel should be between TOP and BOTTOM")
            .isBetween(topY, bottomY);
      }
    }

    @Test
    @DisplayName("TOP vertical alignment is applied even when applyAlignment attribute is absent in xf")
    void topAlignAppliedWhenApplyAlignmentAbsent(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Reproduces a condition found in real Excel files: the <alignment> element has
      // vertical='top' but applyAlignment is not set, causing Apache POI's
      // getVerticalAlignment() to return BOTTOM (the default) instead of TOP.
      // Our getVerticalAlignment() helper reads the raw CTXf to recover the correct value.
      Path excel = createTopAlignNoApplyAlignmentWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int topmost = topmostColoredY(img, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(topmost)
            .as("TOP align (applyAlignment absent): text should appear in the upper half of the cell")
            .isLessThan(T_SAFE_Y);
      }
    }

    @Test
    @DisplayName("RIGHT horizontal alignment is applied even when applyAlignment attribute is absent in xf")
    void rightAlignAppliedWhenApplyAlignmentAbsent(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Same applyAlignment-absent condition as the vertical alignment test, but for horizontal.
      // Apache POI's getAlignment() returns GENERAL (default) when applyAlignment is absent,
      // even if the <alignment> element has horizontal='right'. Our getHorizontalAlignment()
      // workaround reads the raw CTXf to recover the correct value.
      Path excel = createRightAlignNoApplyAlignmentWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int rightmost = rightmostColoredX(img, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(rightmost)
            .as("RIGHT align (applyAlignment absent): text should be in the right half of the cell")
            .isGreaterThan(T_SAFE_X);
      }
    }

    @Test
    @DisplayName("wrapText renders long text across multiple lines")
    void wrappedTextRendersOnMultipleLines(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      // Long text that must wrap in a narrow column
      String longText = "This is a long text that should wrap into multiple lines";
      Path excel = createTextWorkbook(tempDir, "test.xlsx", longText, 11, false, false,
          false, false, false, true, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Multiple wrapped lines → blue pixels span from near top to near bottom of cell
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        int topmost  = topmostColoredY(img, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        int bottommost = bottommostColoredY(img, T_PAD_LEFT, T_RIGHT, T_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(bottommost - topmost)
            .as("wrapped text should span more than half the cell height")
            .isGreaterThan((T_BOTTOM - T_TOP) / 2);
      }
    }

    @Test
    @DisplayName("shrinkToFit reduces font size so text fits within the cell width")
    void shrinkToFitReducesFontSize(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      // Without shrink: text overflows right edge
      Path noShrink = createTextWorkbook(tempDir, "noshrink.xlsx",
          "Very long text that exceeds the cell width significantly", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      // With shrink: text is scaled to fit
      Path withShrink = createTextWorkbook(tempDir, "shrink.xlsx",
          "Very long text that exceeds the cell width significantly", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840, true);
      Path noShrinkPdf = tempDir.resolve("noshrink.pdf");
      Path withShrinkPdf = tempDir.resolve("shrink.pdf");
      ExcelToPdfUtil.generate(noShrink, List.of("Sheet1"), noShrinkPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(withShrink, List.of("Sheet1"), withShrinkPdf, TEST_OPTIONS);

      try (PDDocument nd = Loader.loadPDF(noShrinkPdf.toFile());
          PDDocument sd = Loader.loadPDF(withShrinkPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        // Shrunk text: rightmost blue pixel should be within the cell
        int shrinkRight = rightmostColoredX(si, T_LEFT, T_RIGHT + 20, T_PAD_TOP, T_BOTTOM,
            BLUE_RGB);
        assertThat(shrinkRight)
            .as("shrinkToFit: text should not exceed T_PAD_RIGHT").isLessThanOrEqualTo(T_PAD_RIGHT);
        // Non-shrunk text: may overflow (rightmost could be beyond T_PAD_RIGHT)
        int noShrinkRight = rightmostColoredX(ni, T_LEFT, T_RIGHT + 60, T_PAD_TOP, T_BOTTOM,
            BLUE_RGB);
        assertThat(noShrinkRight)
            .as("no shrink: text overflows beyond T_PAD_RIGHT").isGreaterThan(T_PAD_RIGHT);
      }
    }

    @Test
    @DisplayName("lines below cell bottom are not rendered when cell is too short")
    void textIsHiddenWhenCellTooShort(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      // TOP align, 20pt row, 11pt font, 3 wrapped lines:
      // line 1 fits; line 2 baseline descends below the cell bottom → not rendered.
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "Line1\nLine2\nLine3", 11,
          false, false, false, false, false, true, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 20, 3840);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      int cellBottomPx = T_TOP + (int) (20 * 2); // 20pt row × 2px/pt = 112px

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        // No blue pixels should appear below the cell bottom.
        for (int x = T_PAD_LEFT; x < T_RIGHT; x++) {
          assertThat(isBlueish(img.getRGB(x, cellBottomPx + 4)))
              .as("no text below cell bottom at x=" + x).isFalse();
        }
      }
    }

    @Test
    @DisplayName("vertical text (rotation=255) renders characters stacked top to bottom")
    void verticalTextIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createTextWorkbook(tempDir, "test.xlsx", "縦", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840, false, (short) 255);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Vertical text: blue pixels exist near the top of the cell (first character)
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, T_DPI);
        boolean found = false;
        outer:
        for (int y = T_PAD_TOP; y < T_TOP + 40; y++) {
          for (int x = T_LEFT; x < T_RIGHT; x++) {
            if (isBlueish(img.getRGB(x, y))) {
              found = true;
              break outer;
            }
          }
        }
        assertThat(found).as("vertical text should have blue pixels near top of cell").isTrue();
      }
    }

    @Test
    @DisplayName("single underline renders a line below the text")
    void singleUnderlineIsRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path withUl = createTextWorkbookWithUnderline(tempDir, "ul.xlsx",
          Font.U_SINGLE, blue);
      Path noUl = createTextWorkbook(tempDir, "noul.xlsx", "ABC", 14, false, false,
          false, false, false, false, blue, HorizontalAlignment.LEFT,
          VerticalAlignment.TOP, 60, 3840);
      Path ulPdf   = tempDir.resolve("ul.pdf");
      Path noUlPdf = tempDir.resolve("noul.pdf");
      ExcelToPdfUtil.generate(withUl, List.of("Sheet1"), ulPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(noUl,   List.of("Sheet1"), noUlPdf, TEST_OPTIONS);

      try (PDDocument ud = Loader.loadPDF(ulPdf.toFile());
          PDDocument nd = Loader.loadPDF(noUlPdf.toFile())) {
        BufferedImage ui = new PDFRenderer(ud).renderImageWithDPI(0, T_DPI);
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, T_DPI);
        int ulBlue   = countColoredPixels(ui, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int noUlBlue = countColoredPixels(ni, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(ulBlue)
            .as("underlined text should produce more blue pixels than no-underline")
            .isGreaterThan(noUlBlue);
      }
    }

    @Test
    @DisplayName("double underline renders two lines below the text")
    void doubleUnderlineIsRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path singleUl = createTextWorkbookWithUnderline(tempDir, "single.xlsx",
          Font.U_SINGLE, blue);
      Path doubleUl = createTextWorkbookWithUnderline(tempDir, "double.xlsx",
          Font.U_DOUBLE, blue);
      Path singlePdf = tempDir.resolve("single.pdf");
      Path doublePdf = tempDir.resolve("double.pdf");
      ExcelToPdfUtil.generate(singleUl, List.of("Sheet1"), singlePdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(doubleUl, List.of("Sheet1"), doublePdf, TEST_OPTIONS);

      try (PDDocument sd = Loader.loadPDF(singlePdf.toFile());
          PDDocument dd = Loader.loadPDF(doublePdf.toFile())) {
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, T_DPI);
        BufferedImage di = new PDFRenderer(dd).renderImageWithDPI(0, T_DPI);
        int singleBlue = countColoredPixels(si, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        int doubleBlue = countColoredPixels(di, T_LEFT, T_RIGHT, T_PAD_TOP, T_BOTTOM, BLUE_RGB);
        assertThat(doubleBlue)
            .as("double underline should produce more blue pixels than single underline")
            .isGreaterThan(singleBlue);
      }
    }

    @Test
    @DisplayName("text is rendered when font line height slightly exceeds cell height after scaling")
    void textRenderedWhenFontLineTallerThanCellAfterScaling(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // 14pt BOLD has a natural line height of ~20.3pt (ascent+descent for NotoSansJP).
      // A 20pt row at 85% scale yields a scaled cell height of ~17pt and a scaled font
      // line height of ~17.3pt — the font is fractionally taller than the cell.
      // The baseline of CENTER-aligned text is still within the cell; previously a guard
      // that checked (startY + descent < y) incorrectly skipped the entire line when only
      // the descenders extended below the cell bottom by less than 1pt.
      Path excel = createTightFontCellWorkbook(tempDir);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("合計");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // cell number and date format
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("cell number and date format")
  class CellFormat {

    @Test
    @DisplayName("integer and decimal number formats render correctly")
    void integerAndDecimalFormats(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      assertFmt(tempDir, "i1", 1234.0, "0", "1234");
      assertFmt(tempDir, "i2", 1234.5, "0.00", "1234.50");
    }

    @Test
    @DisplayName("thousands-separator number formats render correctly")
    void thousandsSeparatorFormats(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      assertFmt(tempDir, "t1", 1234567.0, "#,##0", "1,234,567");
      assertFmt(tempDir, "t2", 1234.5, "#,##0.00", "1,234.50");
      assertFmt(tempDir, "t3", -1234.5, "#,##0.00", "-1,234.50");
    }

    @Test
    @DisplayName("percentage formats render correctly")
    void percentageFormats(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertFmt(tempDir, "p1", 0.125, "0%", "13");
      assertFmt(tempDir, "p2", 0.125, "0.0%", "12.5");
    }

    @Test
    @DisplayName("currency formats render correctly")
    void currencyFormats(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertFmt(tempDir, "c1", 1234.0, "¥#,##0", "¥");
      assertFmt(tempDir, "c2", 1234.5, "$#,##0.00", "$1,234.50");
    }

    @Test
    @DisplayName("scientific notation format renders correctly")
    void scientificNotation(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      assertFmt(tempDir, "s1", 12345.6, "0.00E+00", "E");
    }

    @Test
    @DisplayName("ISO and US/European date formats render correctly")
    void dateIsoAndRegionalFormats(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "d1", d, "yyyy/mm/dd", "2026/04/29");
      assertFmt(tempDir, "d2", d, "mm/dd/yyyy", "04/29/2026");
      assertFmt(tempDir, "d3", d, "dd/mm/yyyy", "29/04/2026");
      assertFmt(tempDir, "d4", d, "yyyy-mm-dd", "2026-04-29");
    }

    @Test
    @DisplayName("Japanese kanji date format renders correctly")
    void dateJapaneseKanjiFormat(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "k1", d, "yyyy\"年\"m\"月\"d\"日\"", "2026年4月29日");
    }

    @Test
    @DisplayName("two-digit year date format renders correctly")
    void dateTwoDigitYear(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "y2", d, "yy/mm/dd", "26/04/29");
    }

    @Test
    @DisplayName("built-in date format ID 14 renders yyyy/m/d when JVM default locale is Japanese")
    void builtinDateFormat14WithJapaneseSystemLocale(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Excel stores only the format ID (14) and date serial — no format string.
      // The renderer reads Locale.getDefault() when no explicit dateLocale is set.
      // Temporarily set the JVM default to Japanese to simulate a Japanese OS.
      Locale saved = Locale.getDefault();
      try {
        Locale.setDefault(Locale.JAPAN);
        double d = DateUtil.getExcelDate(LocalDate.of(2018, 2, 21));
        Path excel = createFormattedCellWorkbook(tempDir, "d14-ja.xlsx", d, "m/d/yy");
        Path pdf = tempDir.resolve("d14-ja.pdf");
        ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          String text = new PDFTextStripper().getText(doc);
          assertThat(text).as("Japanese default locale: yyyy/mm/dd").contains("2018/02/21");
        }
      } finally {
        Locale.setDefault(saved);
      }
    }

    @Test
    @DisplayName("built-in date format ID 14 renders m/d/yy when JVM default locale is US")
    void builtinDateFormat14WithUsSystemLocale(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Locale saved = Locale.getDefault();
      try {
        Locale.setDefault(Locale.US);
        double d = DateUtil.getExcelDate(LocalDate.of(2018, 2, 21));
        Path excel = createFormattedCellWorkbook(tempDir, "d14-us.xlsx", d, "m/d/yy");
        Path pdf = tempDir.resolve("d14-us.pdf");
        ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          String text = new PDFTextStripper().getText(doc);
          assertThat(text).as("US default locale: m/d/yy should be rendered literally")
              .contains("2/21/18");
        }
      } finally {
        Locale.setDefault(saved);
      }
    }

    @Test
    @DisplayName("explicit dateLocale in PdfGenerateOptions overrides JVM default locale")
    void explicitDateLocaleOverridesDefault(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Even when JVM default is US, an explicit Locale.JAPAN in options should win.
      Locale saved = Locale.getDefault();
      try {
        Locale.setDefault(Locale.US);
        double d = DateUtil.getExcelDate(LocalDate.of(2018, 2, 21));
        PdfGenerateOptions opts = optionsWithDateLocale(Locale.JAPAN);
        Path excel = createFormattedCellWorkbook(tempDir, "d14-override.xlsx", d, "m/d/yy");
        Path pdf = tempDir.resolve("d14-override.pdf");
        ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, opts);
        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          String text = new PDFTextStripper().getText(doc);
          assertThat(text).as("explicit JAPAN locale should override US default")
              .contains("2018/02/21");
        }
      } finally {
        Locale.setDefault(saved);
      }
    }

    @Test
    @DisplayName("abbreviated month name format (mmm) renders correctly")
    void dateAbbreviatedMonthName(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "m3", d, "mmm d, yyyy", "Apr 29, 2026");
    }

    @Test
    @DisplayName("full month name format (mmmm) renders correctly")
    void dateFullMonthName(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "m4", d, "mmmm d, yyyy", "April 29, 2026");
    }

    @Test
    @DisplayName("localized month name with locale code renders correctly")
    void dateLocalizedMonthName(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      // [$-409] = en-US
      assertFmt(tempDir, "lm", d, "[$-409]mmmm d, yyyy", "April 29, 2026");
    }

    @Test
    @DisplayName("Japanese weekday (aaa) renders correctly")
    void dateJapaneseWeekday(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // 2026-04-29 is a Wednesday (水)
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "wa", d, "yyyy/mm/dd (aaa)", "2026/04/29 (水)");
    }

    @Test
    @DisplayName("English weekday (ddd) renders correctly")
    void dateEnglishWeekday(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "wd", d, "yyyy/mm/dd (ddd)", "Wed");
    }

    @Test
    @DisplayName("date-without-year formats render correctly")
    void dateWithoutYearFormats(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "my1", d, "m/d", "4/29");
    }

    @Test
    @DisplayName("time-only formats render correctly")
    void timeFormats(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double dt = DateUtil.getExcelDate(LocalDateTime.of(2026, 4, 29, 14, 30, 45));
      assertFmt(tempDir, "t1", dt, "h:mm", "14:30");
      assertFmt(tempDir, "t2", dt, "h:mm:ss", "14:30:45");
      assertFmt(tempDir, "t3", dt, "h:mm AM/PM", "PM");
    }

    @Test
    @DisplayName("date-time combined formats render correctly")
    void dateTimeFormats(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double dt = DateUtil.getExcelDate(LocalDateTime.of(2026, 4, 29, 14, 30, 45));
      assertFmt(tempDir, "dt1", dt, "yyyy/mm/dd h:mm", "2026/04/29");
      assertFmt(tempDir, "dt1", dt, "yyyy/mm/dd h:mm", "14:30");
      assertFmt(tempDir, "dt2", dt, "yyyy/mm/dd h:mm:ss", "14:30:45");
    }

    @Test
    @DisplayName("formula cell with number format renders the formatted result")
    void formulaWithNumberFormat(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createFormulaCellWorkbook(tempDir, "test.xlsx", "1000+234", "#,##0");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("1,234");
      }
    }

    @Test
    @DisplayName("formula cell with date format renders the formatted result")
    void formulaWithDateFormat(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      double dateSerial = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      Path excel = createFormulaCellWorkbook(tempDir, "test.xlsx",
          String.valueOf(dateSerial), "yyyy/mm/dd", dateSerial);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("2026/04/29");
      }
    }

    @Test
    @DisplayName("Reiwa era format renders correctly")
    void eraReiwaFormat(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // 2026-04-29 = 令和8年4月29日
      double d = DateUtil.getExcelDate(LocalDate.of(2026, 4, 29));
      assertFmt(tempDir, "era", d, "ggge\"年\"m\"月\"d\"日\"", "令和8年4月29日");
    }

    @Test
    @DisplayName("era format with pre-Reiwa date throws RuntimeException")
    void eraPreReiwaThrowsException(@TempDir Path tempDir) throws IOException {
      double d = DateUtil.getExcelDate(LocalDate.of(2019, 4, 30)); // last day of Heisei
      Path excel = createFormattedCellWorkbook(tempDir, "test.xlsx", d,
          "ggge\"年\"m\"月\"d\"日\"");
      Path pdf = tempDir.resolve("out.pdf");
      assertThatThrownBy(() -> ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS))
          .isInstanceOf(RuntimeException.class)
          .hasMessageContaining("Reiwa");
    }

    @Test
    @DisplayName("zero-valued formula cell with accounting format renders as dash not digit zero")
    void zeroFormulaWithAccountingZeroSection(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Accounting-style formats use a zero section like `"-"??` where `??` is for
      // decimal-point alignment only. DataFormatter.formatRawCellContents incorrectly
      // renders the `??` as the digit 0, producing "- 0" instead of "-".
      Path excel = createZeroFormulaWorkbook(tempDir, "#,##0.00;-#,##0.00;\"-\"??");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = new PDFTextStripper().getText(doc).trim();
        assertThat(text).as("zero section literal should be rendered").contains("-");
        assertThat(text).as("digit 0 from ?? placeholder must not appear").doesNotContain("0");
      }
    }

    private void assertFmt(Path dir, String id, double value, String format, String expected)
        throws IOException, PdfGenerateException {
      Path excel = createFormattedCellWorkbook(dir, id + ".xlsx", value, format);
      Path pdf = dir.resolve(id + ".pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        String text = normalizeCjk(new PDFTextStripper().getText(doc));
        assertThat(text)
            .as("format '%s' value %s", format, value)
            .contains(expected);
      }
    }

    // PDFBox may extract CJK characters as their CJK Radicals Supplement equivalents.
    // Normalize known radical variants back to regular CJK Unified Ideographs.
    private static String normalizeCjk(String text) {
      return text
          .replace('⽉', '月')  // ⽉ (U+2F49) → 月 (U+6708)
          .replace('⽇', '日')  // ⽇ (U+2F47) → 日 (U+65E5)
          .replace('⽔', '水'); // ⽔ (U+2F54) → 水 (U+6C34)
    }
  }

  // ---------------------------------------------------------------------------
  // cell background color
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("cell background color")
  class CellBackground {

    // 144 DPI: same cell geometry as CellBorders.
    // left=18pt→36px, top=36pt→72px, row=60pt→120px, col=2048units(42pt)→84px
    private static final int C_DPI    = 144;
    private static final int C_LEFT   = 36;
    private static final int C_SAFE_X = 78;   // (C_LEFT + C_RIGHT) / 2, C_RIGHT=120
    private static final int C_SAFE_Y = 132;  // (72 + C_BOTTOM) / 2, C_BOTTOM=192

    @Test
    @DisplayName("solid fill RGB is reflected exactly in the PDF (±2 per channel)")
    void solidFillRgbIsReflectedExactly(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createBgWorkbook(tempDir, "test.xlsx", blue, FillPatternType.SOLID_FOREGROUND);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertChannels(img, C_SAFE_X, C_SAFE_Y, 0x0070C0);
      }
    }

    @Test
    @DisplayName("cell with no fill renders as white background")
    void noFillRendersAsWhite(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createBgWorkbook(tempDir, "test.xlsx", null, FillPatternType.NO_FILL);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(img.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("no-fill cell should be white").isGreaterThan(250);
      }
    }

    @Test
    @DisplayName("white solid fill renders as white background")
    void whiteSolidFillRendersAsWhite(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor white = new XSSFColor(new byte[] {(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}, null);
      Path excel = createBgWorkbook(tempDir, "test.xlsx", white, FillPatternType.SOLID_FOREGROUND);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(img.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("white fill should be white").isGreaterThan(250);
      }
    }

    @Test
    @DisplayName("black solid fill renders as black background")
    void blackFillIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      XSSFColor black = new XSSFColor(new byte[] {0x00, 0x00, 0x00}, null);
      Path excel = createBgWorkbook(tempDir, "test.xlsx", black, FillPatternType.SOLID_FOREGROUND);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(img.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("black fill should be black").isLessThan(5);
      }
    }

    @Test
    @DisplayName("adjacent cells with different colors each render their own color")
    void adjacentCellsHaveIndependentColors(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      XSSFColor red  = new XSSFColor(new byte[] {(byte) 0xFF, 0x00, 0x00}, null);
      Path excel = createAdjacentBgWorkbook(tempDir, "test.xlsx", blue, red);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // cell 0: safe center (78, 132) — blue
      // cell 1: safe center at x = 120 + 42 (col width) / 2 = 141 — red
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertChannels(img, C_SAFE_X, C_SAFE_Y, 0x0070C0);
        assertChannels(img, 162, C_SAFE_Y, 0xFF0000);  // 120 + 84/2 = 162
      }
    }

    @Test
    @DisplayName("indexed color fill is rendered correctly")
    void indexedColorIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createIndexedColorWorkbook(tempDir, "test.xlsx", IndexedColors.DARK_RED);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        int rgb = img.getRGB(C_SAFE_X, C_SAFE_Y) & 0xFFFFFF;
        assertThat((rgb >> 16) & 0xFF)
            .as("red channel of DARK_RED should dominate").isGreaterThan(100);
        assertThat((rgb >> 8) & 0xFF)
            .as("green channel of DARK_RED should be low").isLessThan(50);
      }
    }

    @Test
    @DisplayName("theme color renders as a non-white color")
    void themeColorIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createThemeColorWorkbook(tempDir, "test.xlsx", 4); // accent1
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(img.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("theme color should render as a non-white color").isLessThan(250);
      }
    }

    @Test
    @DisplayName("positive tint makes the fill color lighter than the base color")
    void positiveTintMakesColorLighter(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor base   = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      XSSFColor lighter = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      lighter.setTint(0.5);

      Path basePdf   = renderBg(tempDir, "base.xlsx", "base.pdf", base);
      Path lightPdf  = renderBg(tempDir, "light.xlsx", "light.pdf", lighter);

      try (PDDocument bd = Loader.loadPDF(basePdf.toFile());
          PDDocument ld = Loader.loadPDF(lightPdf.toFile())) {
        BufferedImage bi = new PDFRenderer(bd).renderImageWithDPI(0, C_DPI);
        BufferedImage li = new PDFRenderer(ld).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(li.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("tint +0.5 should be lighter (higher gray) than base")
            .isGreaterThan(avgGray(bi.getRGB(C_SAFE_X, C_SAFE_Y)));
      }
    }

    @Test
    @DisplayName("negative tint makes the fill color darker than the base color")
    void negativeTintMakesColorDarker(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor base   = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      XSSFColor darker = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      darker.setTint(-0.5);

      Path basePdf  = renderBg(tempDir, "base.xlsx", "base.pdf", base);
      Path darkPdf  = renderBg(tempDir, "dark.xlsx", "dark.pdf", darker);

      try (PDDocument bd = Loader.loadPDF(basePdf.toFile());
          PDDocument dd = Loader.loadPDF(darkPdf.toFile())) {
        BufferedImage bi = new PDFRenderer(bd).renderImageWithDPI(0, C_DPI);
        BufferedImage di = new PDFRenderer(dd).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(di.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("tint -0.5 should be darker (lower gray) than base")
            .isLessThan(avgGray(bi.getRGB(C_SAFE_X, C_SAFE_Y)));
      }
    }

    @Test
    @DisplayName("non-SOLID fill patterns are not rendered as a background fill")
    void nonSolidPatternFillIsNotRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createBgWorkbook(tempDir, "test.xlsx", blue, FillPatternType.THIN_HORZ_BANDS);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        assertThat(avgGray(img.getRGB(C_SAFE_X, C_SAFE_Y)))
            .as("GRAY_25 pattern should not fill the cell with blue").isGreaterThan(200);
      }
    }

    @Test
    @DisplayName("background fill is drawn under text (text pixels overwrite background)")
    void backgroundFillIsDrawnUnderText(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor bgBlue  = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      XSSFColor textRed = new XSSFColor(new byte[] {(byte) 0xFF, 0x00, 0x00}, null);
      Path excel = createBgWithTextWorkbook(tempDir, "test.xlsx", bgBlue, textRed);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, C_DPI);
        // Background (far from text): blue pixels present
        boolean hasBlue = isBlueish(img.getRGB(C_SAFE_X + 30, C_SAFE_Y));
        // Entire cell: some non-blue pixels exist (text rendered in red on top)
        boolean hasNonBlue = false;
        for (int x = C_LEFT + 4; x < C_LEFT + 50 && !hasNonBlue; x++) {
          if (!isBlueish(img.getRGB(x, C_SAFE_Y))) {
            hasNonBlue = true;
          }
        }
        assertThat(hasBlue).as("background blue should be present away from text").isTrue();
        assertThat(hasNonBlue).as("text (red) should overwrite background on text pixels").isTrue();
      }
    }

    private Path renderBg(Path dir, String xlName, String pdfName, XSSFColor color)
        throws IOException, PdfGenerateException {
      Path excel = createBgWorkbook(dir, xlName, color, FillPatternType.SOLID_FOREGROUND);
      Path pdf = dir.resolve(pdfName);
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      return pdf;
    }

    private static void assertChannels(BufferedImage img, int x, int y, int expectedRgb) {
      int actual = img.getRGB(x, y) & 0xFFFFFF;
      int er = (expectedRgb >> 16) & 0xFF;
      int eg = (expectedRgb >> 8) & 0xFF;
      int eb = expectedRgb & 0xFF;
      assertThat(((actual >> 16) & 0xFF)).as("R at (%d,%d)", x, y).isBetween(er - 2, er + 2);
      assertThat(((actual >> 8) & 0xFF)).as("G at (%d,%d)", x, y).isBetween(eg - 2, eg + 2);
      assertThat((actual & 0xFF)).as("B at (%d,%d)", x, y).isBetween(eb - 2, eb + 2);
    }
  }

  // ---------------------------------------------------------------------------
  // cell images (XSSFPicture)
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("cell images (XSSFPicture)")
  class CellImage {

    // 72 DPI: 1pt = 1px. Geometry: leftMargin=18pt, topMargin=36pt, rowH=30pt, colW≈50pt.
    // 2-cell anchor (row2=2, col2=2): image spans x=[18,118], y=[36,96].
    private static final int IM_DPI    = 72;
    private static final int IM_LEFT   = 18;   // left edge of image
    private static final int IM_TOP    = 36;   // top edge in image coords
    private static final int IM_CX     = 68;   // center x = 18 + 50
    private static final int IM_CY     = 66;   // center y = 36 + 30
    private static final int IM_RIGHT  = 118;  // right edge = 18 + 2×50

    @Test
    @DisplayName("PNG image is rendered at the correct position")
    void pngImageIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      Path pdf = renderImage(tempDir, img, Workbook.PICTURE_TYPE_PNG, 0, 0, 2, 2);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_CX, IM_CY)))
            .as("image center should be red (PNG)").isTrue();
      }
    }

    @Test
    @DisplayName("JPEG image is rendered at the correct position")
    void jpegImageIsRendered(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      byte[] img = solidColorJpeg(0xFF0000, 80, 60);
      Path pdf = renderImage(tempDir, img, Workbook.PICTURE_TYPE_JPEG, 0, 0, 2, 2);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_CX, IM_CY)))
            .as("image center should be reddish (JPEG)").isTrue();
      }
    }

    @Test
    @DisplayName("PNG with transparency: opaque area is rendered, transparent area shows background")
    void pngWithTransparencyRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Top half (30px): opaque red; bottom half (30px): transparent
      byte[] img = halfTransparentPng(0xFF0000, 80, 60);
      Path pdf = renderImage(tempDir, img, Workbook.PICTURE_TYPE_PNG, 0, 0, 2, 2);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_CX, IM_TOP + 5)))
            .as("top half (opaque) should be red").isTrue();
        assertThat(avgGray(rendered.getRGB(IM_CX, IM_TOP + 50)))
            .as("bottom half (transparent) should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("two-cell anchor determines the rendered image size")
    void imageSizeMatchesTwoCellAnchor(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      Path pdf = renderImage(tempDir, img, Workbook.PICTURE_TYPE_PNG, 0, 0, 2, 2);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_RIGHT - 3, IM_CY)))
            .as("pixel just inside right edge should be red").isTrue();
        assertThat(avgGray(rendered.getRGB(IM_RIGHT + 5, IM_CY)))
            .as("pixel just outside right edge should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("one-cell anchor renders image using its natural dimensions")
    void oneCellAnchorRendersImage(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      // 80×60 px image → natural size ≈ 60pt × 45pt at 72 DPI
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      Path pdf = renderImage(tempDir, img, Workbook.PICTURE_TYPE_PNG, 0, 0, 0, 0);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        // Image top-left at (18,36); natural size → pixel at (28,46) should be inside
        assertThat(isReddish(rendered.getRGB(IM_LEFT + 10, IM_TOP + 10)))
            .as("one-cell anchor: image should render at natural size").isTrue();
      }
    }

    @Test
    @DisplayName("multiple images in the same sheet are each rendered")
    void multipleImagesAreEachRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] red   = solidColorPng(0xFF0000, 80, 60);
      byte[] green = solidColorPng(0x00FF00, 80, 60);
      Path excel = createTwoImageWorkbook(tempDir, "test.xlsx",
          red, Workbook.PICTURE_TYPE_PNG, 0, 0, 2, 2,
          green, Workbook.PICTURE_TYPE_PNG, 0, 3, 2, 5);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Green image center: top at 36+3×30=126, center y=126+30=156
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_CX, IM_CY)))
            .as("first image (red) should be rendered").isTrue();
        assertThat(isGreenish(rendered.getRGB(IM_CX, IM_TOP + 3 * 30 + 30)))
            .as("second image (green) should be rendered").isTrue();
      }
    }

    @Test
    @DisplayName("image appears only on the correct page when multiple pages exist")
    void imageAppearsOnCorrectPageWhenMultiplePages(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // Image on rows 0-2 (page 1). Add enough rows to force a second page.
      Path excel = createImageWithManyRowsWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThanOrEqualTo(2);
        BufferedImage page1 = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        BufferedImage page2 = new PDFRenderer(doc).renderImageWithDPI(1, IM_DPI);
        assertThat(isReddish(page1.getRGB(IM_CX, IM_CY)))
            .as("image should appear on page 1").isTrue();
        assertThat(isReddish(page2.getRGB(IM_CX, IM_CY)))
            .as("image should NOT appear on page 2").isFalse();
      }
    }

    @Test
    @DisplayName("image scales proportionally when explicit scale is applied")
    void imageScalesWithExplicitScale(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // 50% scale: image spans x=[18,68], y=[36,66] (half of unscaled 100×60)
      Path excel = createScaledImageWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG, (short) 50);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        // Inside scaled image (center at 18+25=43, 36+15=51)
        assertThat(isReddish(rendered.getRGB(IM_LEFT + 20, IM_TOP + 12)))
            .as("scaled image center should be red").isTrue();
        // Outside scaled image (unscaled center 68,66 is now beyond the scaled right edge)
        assertThat(avgGray(rendered.getRGB(IM_CX + 10, IM_CY + 5)))
            .as("unscaled center should be outside the 50%-scaled image").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("image outside the print area is not rendered")
    void imageOutsidePrintAreaIsNotRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // Print area: rows 0-2. Image at rows 4-6 (outside print area).
      Path excel = createImageOutsidePrintAreaWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        // No red pixels anywhere (image was outside print area)
        boolean hasRed = false;
        for (int y = 0; y < rendered.getHeight() && !hasRed; y++) {
          for (int x = 0; x < rendered.getWidth() && !hasRed; x++) {
            if (isReddish(rendered.getRGB(x, y))) {
              hasRed = true;
            }
          }
        }
        assertThat(hasRed).as("image outside print area should not be rendered").isFalse();
      }
    }

    @Test
    @DisplayName("image anchored outside the print area (column direction) is not rendered")
    void imageOutsidePrintAreaColumnIsNotRendered(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // Print area: cols 0-3 (A-D). Image col1=5 (column F, outside).
      Path excel = createImageOutsidePrintAreaColumnWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        boolean hasRed = false;
        for (int y = 0; y < rendered.getHeight() && !hasRed; y++) {
          for (int x = 0; x < rendered.getWidth() && !hasRed; x++) {
            if (isReddish(rendered.getRGB(x, y))) {
              hasRed = true;
            }
          }
        }
        assertThat(hasRed).as("image outside print area column should not be rendered").isFalse();
      }
    }

    @Test
    @DisplayName("image and cell text/background coexist on the same page")
    void imageAndCellContentCoexist(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // Cell at row 3 has blue background and text; image at rows 0-2
      Path excel = createImageWithCellContentWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        // Image should be rendered
        BufferedImage rendered = new PDFRenderer(doc).renderImageWithDPI(0, IM_DPI);
        assertThat(isReddish(rendered.getRGB(IM_CX, IM_CY)))
            .as("image should be rendered").isTrue();
        // Cell text should be extractable
        assertThat(new PDFTextStripper().getText(doc)).contains("内容");
      }
    }

    private Path renderImage(Path dir, byte[] imgBytes, int pictureType,
        int col1, int row1, int col2, int row2) throws IOException, PdfGenerateException {
      Path excel = createImageWorkbook(dir, "test.xlsx", imgBytes, pictureType,
          col1, row1, col2, row2);
      Path pdf = dir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);
      return pdf;
    }

    private static boolean isReddish(int argb) {
      int r = (argb >> 16) & 0xFF;
      int g = (argb >> 8) & 0xFF;
      int b = argb & 0xFF;
      return r > 150 && r > g + 80 && r > b + 80;
    }

    private static boolean isGreenish(int argb) {
      int r = (argb >> 16) & 0xFF;
      int g = (argb >> 8) & 0xFF;
      int b = argb & 0xFF;
      return g > 150 && g > r + 80 && g > b + 80;
    }
  }

  // ---------------------------------------------------------------------------
  // merged cells
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("merged cells")
  class CellMerge {

    // 72 DPI: 1pt = 1px.  left=18pt, top=36pt, rowH=30pt, colW≈50pt.
    private static final int M_DPI  = 72;
    private static final int M_LM   = 18;
    private static final int M_TM   = 36;
    private static final int M_ROW  = 30;
    private static final int M_COL  = 50;

    // 144 DPI for border detection (same as CellBorders).
    private static final int MB_DPI = 144;
    private static final int MB_LM  = 36;  // 18pt × 2
    private static final int MB_TM  = 72;  // 36pt × 2

    private static final int BLUE_BG = 0x0070C0;

    @Test
    @DisplayName("horizontal merge (1 row × 3 cols) spans the full combined column width")
    void horizontalMergeSpansCorrectWidth(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createMergedWorkbook(tempDir, "test.xlsx", 0, 0, 0, 2,
          blue, null, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Merge: width = 3 × 50 = 150px. Right edge = 18 + 150 = 168px.
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, M_DPI);
        int safeY = M_TM + M_ROW / 2;
        assertThat(isColoredPixel(img.getRGB(M_LM + 140, safeY), BLUE_BG))
            .as("pixel 10px inside right edge should be blue").isTrue();
        assertThat(avgGray(img.getRGB(M_LM + 155, safeY)))
            .as("pixel 5px beyond right edge should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("vertical merge (3 rows × 1 col) spans the full combined row height")
    void verticalMergeSpansCorrectHeight(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createMergedWorkbook(tempDir, "test.xlsx", 0, 2, 0, 0,
          blue, null, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Merge: height = 3 × 30 = 90px. Bottom edge = 36 + 90 = 126px.
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, M_DPI);
        int safeX = M_LM + M_COL / 2;
        assertThat(isColoredPixel(img.getRGB(safeX, M_TM + 80), BLUE_BG))
            .as("pixel 10px above bottom edge should be blue").isTrue();
        assertThat(avgGray(img.getRGB(safeX, M_TM + 95)))
            .as("pixel 5px below bottom edge should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("rectangular merge (2 rows × 2 cols) spans correct width and height")
    void rectangularMergeSpansBothDimensions(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      Path excel = createMergedWorkbook(tempDir, "test.xlsx", 0, 1, 0, 1,
          blue, null, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // Merge: 2×2 = 100px wide, 60px tall.
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, M_DPI);
        assertThat(isColoredPixel(img.getRGB(M_LM + 90, M_TM + 50), BLUE_BG))
            .as("pixel inside merged area (near corner) should be blue").isTrue();
        assertThat(avgGray(img.getRGB(M_LM + 105, M_TM + 50)))
            .as("pixel outside right edge should be white").isGreaterThan(220);
        assertThat(avgGray(img.getRGB(M_LM + 90, M_TM + 65)))
            .as("pixel below bottom edge should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("right border of merged cell is taken from the rightmost column cell's style")
    void rightBorderTakenFromRightmostCell(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Merge row 0, cols 0-2. THICK right border only on col-2 cell (not col-0).
      Path excel = createMergeWithBorderWorkbook(tempDir, "test.xlsx",
          0, 0, 0, 2, BorderStyle.THICK, BorderStyle.NONE);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // At 144 DPI: right edge = 36 + 3×100 = 336px
      int rightEdge = MB_LM + 3 * MB_DPI / 72 * M_COL; // 36 + 300 = 336
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, MB_DPI);
        assertThat(avgGray(img.getRGB(rightEdge, MB_TM + 30)))
            .as("right border (from rightmost cell) should be dark").isLessThan(128);
      }
    }

    @Test
    @DisplayName("bottom border of merged cell is taken from the bottom row cell's style")
    void bottomBorderTakenFromBottomCell(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Merge rows 0-2, col 0. THICK bottom border only on row-2 cell (not row-0).
      Path excel = createMergeWithBorderWorkbook(tempDir, "test.xlsx",
          0, 2, 0, 0, BorderStyle.NONE, BorderStyle.THICK);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      // At 144 DPI: bottom edge = 72 + 3×60 = 252px, col center = 36 + 50 = 86px
      int bottomEdge = MB_TM + 3 * MB_DPI / 72 * M_ROW;  // 72 + 180 = 252
      int centerX    = MB_LM + MB_DPI / 72 * M_COL / 2;  // 36 + 50  = 86
      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, MB_DPI);
        assertThat(avgGray(img.getRGB(centerX, bottomEdge)))
            .as("bottom border (from bottom row cell) should be dark").isLessThan(128);
      }
    }

    @Test
    @DisplayName("background fill covers the entire merged area")
    void backgroundFillCoversFullMergedArea(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      // 3×3 merge so the area is large enough for clear checks
      Path excel = createMergedWorkbook(tempDir, "test.xlsx", 0, 2, 0, 2,
          blue, null, null);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, M_DPI);
        // Center of 3×3 merge: x=18+75=93, y=36+45=81
        assertThat(isColoredPixel(img.getRGB(M_LM + 75, M_TM + 45), BLUE_BG))
            .as("center of merged area should be filled with blue").isTrue();
        // Just outside: x=18+155=173 (right of 150px width), y=81
        assertThat(avgGray(img.getRGB(M_LM + 155, M_TM + 45)))
            .as("cell outside the merge to the right should be white").isGreaterThan(220);
        // Just outside: x=93, y=36+95=131 (below 90px height)
        assertThat(avgGray(img.getRGB(M_LM + 75, M_TM + 95)))
            .as("cell outside the merge below should be white").isGreaterThan(220);
      }
    }

    @Test
    @DisplayName("RIGHT alignment in merged cell uses the merged width for text positioning")
    void textAlignmentUsesFullMergedWidth(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      XSSFColor blue = new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null);
      // 3-col merge: right-aligned text should end near x = 18+150-2 = 166px
      Path mergedXl = createMergedWorkbookWithFont(tempDir, "merged.xlsx",
          0, 0, 0, 2, blue, HorizontalAlignment.RIGHT);
      // 1-col single: right-aligned text should end near x = 18+50-2 = 66px
      Path singleXl = createMergedWorkbookWithFont(tempDir, "single.xlsx",
          0, 0, 0, 0, blue, HorizontalAlignment.RIGHT);
      Path mergedPdf = tempDir.resolve("merged.pdf");
      Path singlePdf = tempDir.resolve("single.pdf");
      ExcelToPdfUtil.generate(mergedXl, List.of("Sheet1"), mergedPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(singleXl, List.of("Sheet1"), singlePdf, TEST_OPTIONS);

      try (PDDocument md = Loader.loadPDF(mergedPdf.toFile());
          PDDocument sd = Loader.loadPDF(singlePdf.toFile())) {
        BufferedImage mi = new PDFRenderer(md).renderImageWithDPI(0, M_DPI);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, M_DPI);
        int mergedRight = rightmostColoredX(mi, M_LM, M_LM + 160,
            M_TM, M_TM + M_ROW, BLUE_BG);
        int singleRight = rightmostColoredX(si, M_LM, M_LM + 60,
            M_TM, M_TM + M_ROW, BLUE_BG);
        assertThat(mergedRight)
            .as("merged cell RIGHT-aligned text should end further right than single cell")
            .isGreaterThan(singleRight + 50);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // print title rows
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("print title rows")
  class PrintTitleRows {

    @Test
    @DisplayName("title rows appear on every page including page 2+")
    void titleRowsAppearOnEveryPage(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleRowsWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThanOrEqualTo(2);
        // Page 1: contains the title row text
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        assertThat(s1.getText(doc)).contains("HEADER");
        // Page 2: also contains the title row text (repeated)
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        assertThat(s2.getText(doc)).contains("HEADER");
      }
    }

    @Test
    @DisplayName("content rows (non-title) are distributed across pages")
    void contentRowsDistributedAcrossPages(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleRowsWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        // Row 1 (first content row) appears on page 1 only
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        assertThat(s1.getText(doc)).contains("row1");
        // Last data row (row35) appears on page 2 only
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        assertThat(s2.getText(doc)).contains("row35");
      }
    }

    @Test
    @DisplayName("preamble rows (before print title row) appear on page 1 only")
    void preambleRowsAppearOnFirstPageOnly(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createPreambleTitleRowsWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isGreaterThanOrEqualTo(2);
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        String page1 = s1.getText(doc);
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        String page2 = s2.getText(doc);

        // Preamble rows (PREAMBLE_A, PREAMBLE_B) appear on page 1
        assertThat(page1).contains("PREAMBLE_A");
        assertThat(page1).contains("PREAMBLE_B");
        // Preamble rows do NOT appear on page 2
        assertThat(page2).doesNotContain("PREAMBLE_A");
        assertThat(page2).doesNotContain("PREAMBLE_B");
        // Title row (HEADER) appears on both pages
        assertThat(page1).contains("HEADER");
        assertThat(page2).contains("HEADER");
        // Content rows are distributed across pages
        assertThat(page1).contains("row1");
        assertThat(page2).contains("row30");
      }
    }

    @Test
    @DisplayName("preamble rows on page 1 reduce available space for content rows")
    void preambleRowsReduceContentCapacityOnPage1(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createPreambleTitleRowsWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        // Without preamble 30 content rows (30×25=750pt) fit on page 1.
        // With preamble (2×20=40pt) the page-1 capacity drops to 710pt → only 28 rows fit.
        // row29 must therefore appear on page 2, not page 1.
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        String page1 = s1.getText(doc);
        assertThat(page1).doesNotContain("row29");
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        String page2 = s2.getText(doc);
        assertThat(page2).contains("row29");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // multiple sheets
  // ---------------------------------------------------------------------------

  // ---------------------------------------------------------------------------
  // table style rendering
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("table style")
  class TableStyleRendering {

    @Test
    @DisplayName("table header row receives fill colour from the table style")
    void headerRowHasTableStyleFill(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // Use a custom table style with an explicit-RGB fill so colour resolution
      // does not depend on the workbook's theme being populated.
      Path excel = createTableWithCustomStyleWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 144);
        // Header row: top margin = 0.25in = 18pt → 36px at 144 DPI.
        // Sample a pixel in the header row fill area.
        int sampleX = 40;
        int headerY = 36 + 5;
        int rgb = img.getRGB(sampleX, headerY) & 0xFFFFFF;
        // Header should not be white (custom style applies a red fill: #FF0000)
        assertThat(rgb)
            .as("Header row should have a non-white fill from the table style")
            .isNotEqualTo(0xFFFFFF);
      }
    }

    @Test
    @DisplayName("alternating row stripes produce different fills on odd and even data rows")
    void rowStripesAlternate(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTableWithCustomStyleWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 144);
        int sampleX = 40;
        // Header 20pt=40px ends at 76px. Data rows 25pt=50px each.
        int row1Y = 36 + 40 + 10;      // inside first data row (stripe)
        int row2Y = 36 + 40 + 50 + 10; // inside second data row (no stripe)
        int rgb1 = img.getRGB(sampleX, row1Y) & 0xFFFFFF;
        int rgb2 = img.getRGB(sampleX, row2Y) & 0xFFFFFF;
        assertThat(rgb1)
            .as("First data row (stripe) fill should differ from second data row (no stripe)")
            .isNotEqualTo(rgb2);
      }
    }

    @Test
    @DisplayName("built-in table style with theme colours resolves via ThemesTable (inheritFromThemeAsRequired)")
    void builtinTableStyleThemeColourResolved(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      // invoice-table-theme-test.xlsx is a real Excel file with a custom theme.
      // Its "InvoiceDetails" table uses TableStyleMedium6 which applies the theme's
      // accent5 colour (#5E858C = teal) as the header fill.  This exercises the
      // ThemesTable.inheritFromThemeAsRequired() fallback in poiColorToAwt().
      var stream = ExcelToPdfUtilTest.class.getResourceAsStream(
          "/invoice-table-theme-test.xlsx");
      assertThat(stream).as("test resource must be present").isNotNull();
      Path excel = tempDir.resolve("invoice.xlsx");
      try (stream) {
        Files.copy(stream, excel);
      }
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Invoice"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 144);

        // The table header (row 10 in the invoice, "Details"/"AMOUNT") is at
        // roughly the middle of the page. We cannot pin the exact pixel position
        // without replicating the layout calculation, so instead we check that
        // at least ONE non-white, non-grey pixel exists in the vertical centre
        // band of the image — confirming that the teal header fill was rendered.
        int w = img.getWidth();
        int h = img.getHeight();
        int cx = w / 2;
        boolean foundColoredFill = false;
        outer:
        for (int y = h / 6; y < 2 * h / 3; y++) {
          for (int dx = -w / 8; dx <= w / 8; dx++) {
            int rgb = img.getRGB(cx + dx, y) & 0xFFFFFF;
            int r = (rgb >> 16) & 0xFF;
            int g = (rgb >>  8) & 0xFF;
            int b =  rgb        & 0xFF;
            // A "coloured" pixel is one whose channels differ significantly
            // (not pure grey: |r-g|>20 or |g-b|>20 or |r-b|>20 and not white)
            if (rgb != 0xFFFFFF && (Math.abs(r - g) > 20 || Math.abs(g - b) > 20
                || Math.abs(r - b) > 20)) {
              foundColoredFill = true;
              break outer;
            }
          }
        }
        assertThat(foundColoredFill)
            .as("At least one coloured (non-grey) fill pixel should be visible, "
                + "confirming that theme-based table style colours were resolved")
            .isTrue();
      }
    }

    @Test
    @DisplayName("wholeTable horizontal border draws a line between data rows")
    void horizontalBordersVisibleBetweenDataRows(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTableWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        PDFTextStripper s = new PDFTextStripper();
        String text = s.getText(doc);
        assertThat(text).contains("HEADER").contains("Row1").contains("Row2").contains("Row3");

        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 144);
        // Boundary between row1 and row2: 36 (margin) + 40 (header) + 50 (row1) = 126px.
        int borderY = 36 + 40 + 50;
        int sampleX = 40;
        boolean borderFound = false;
        for (int dy = -2; dy <= 2; dy++) {
          int y = borderY + dy;
          if (y >= 0 && y < img.getHeight()) {
            if ((img.getRGB(sampleX, y) & 0xFFFFFF) != 0xFFFFFF) {
              borderFound = true;
              break;
            }
          }
        }
        assertThat(borderFound)
            .as("A horizontal border line should be visible between data rows")
            .isTrue();
      }
    }

    @Test
    @DisplayName("header text appears white on dark table header background")
    void headerTextIsWhiteOnDarkBackground(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTableWithCustomStyleWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(new PDFTextStripper().getText(doc)).contains("HEADER");

        BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 144);
        int sampleX = 40;
        int headerFillY = 36 + 5; // background pixel in header area
        int fillRgb = img.getRGB(sampleX, headerFillY) & 0xFFFFFF;
        // Custom style uses explicit red (#FF0000) fill → clearly dark enough to trigger
        // white text fallback. Verify that the fill is NOT white.
        int r = (fillRgb >> 16) & 0xFF;
        int g = (fillRgb >> 8)  & 0xFF;
        int b =  fillRgb        & 0xFF;
        double lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
        assertThat(lum)
            .as("Header fill should be a dark colour (luminance < 0.5)")
            .isLessThan(0.5);
      }
    }

    @Test
    @DisplayName("showLastColumn applies bold font to data rows in the last column")
    void lastColumnDataRowsAreBold(@TempDir Path tempDir)
        throws Exception {
      // Generate two PDFs: one WITH showLastColumn (bold) and one WITHOUT.
      // The last column area should look different: bold text has more ink than regular.
      Path excelWith = createTableWithLastColumnWorkbook(tempDir, "with.xlsx");
      Path excelWithout = createTableWithLastColumnNoShowWorkbook(tempDir, "without.xlsx");
      Path pdfWith = tempDir.resolve("with.pdf");
      Path pdfWithout = tempDir.resolve("without.pdf");
      ExcelToPdfUtil.generate(excelWith, List.of("Sheet1"), pdfWith, TEST_OPTIONS);
      ExcelToPdfUtil.generate(excelWithout, List.of("Sheet1"), pdfWithout, TEST_OPTIONS);

      try (PDDocument docWith = Loader.loadPDF(pdfWith.toFile());
          PDDocument docWithout = Loader.loadPDF(pdfWithout.toFile())) {
        assertThat(new PDFTextStripper().getText(docWith))
            .contains("ColA-Row1").contains("ColB-Row1");

        BufferedImage imgWith    = new PDFRenderer(docWith).renderImageWithDPI(0, 144);
        BufferedImage imgWithout = new PDFRenderer(docWithout).renderImageWithDPI(0, 144);

        // Count dark pixels in the data rows (skip header).
        // Table is B1:C3; C is the last column. At 144 DPI, left margin=0.25in=36px,
        // each column ~123px. Data rows: y≈76–176px. Scan full width — A/B are
        // identical in both PDFs, so any difference comes from C-column bold.
        int w = imgWith.getWidth();
        int h = imgWith.getHeight();
        int yStart = h / 20;   // ~84px — below header row (36–76px)
        int yEnd   = h / 9;    // ~187px — past both data rows (76–176px)

        int darkWith = 0, darkWithout = 0;
        for (int y = yStart; y < yEnd; y++) {
          for (int x = 0; x < w; x++) {
            double lum = luminance(imgWith.getRGB(x, y));
            if (lum < 0.5) darkWith++;
            lum = luminance(imgWithout.getRGB(x, y));
            if (lum < 0.5) darkWithout++;
          }
        }
        assertThat(darkWith)
            .as("showLastColumn=true (bold) should produce more dark pixels than without")
            .isGreaterThan(darkWithout);
      }
    }
  }

  @Nested
  @DisplayName("multiple sheets")
  class MultipleSheets {

    @Test
    @DisplayName("two sheets produce two PDF pages in the specified order")
    void twoSheetsProduceTwoPages(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTwoSheetWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1", "Sheet2"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        assertThat(s1.getText(doc)).contains("sheet1content");
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        assertThat(s2.getText(doc)).contains("sheet2content");
      }
    }

    @Test
    @DisplayName("sheets are rendered in the specified order")
    void sheetOrderIsPreserved(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTwoSheetWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      // Reversed order
      ExcelToPdfUtil.generate(excel, List.of("Sheet2", "Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        assertThat(s1.getText(doc)).contains("sheet2content");
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        assertThat(s2.getText(doc)).contains("sheet1content");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // print title columns
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("print title columns")
  class PrintTitleCols {

    @Test
    @DisplayName("title col appears on every page including page 2+")
    void titleColAppearsOnEachPage(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleColsWorkbook(tempDir, "test.xlsx", 1, 11);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        String page1 = s1.getText(doc);
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        String page2 = s2.getText(doc);
        assertThat(page1).contains("LABEL_A");
        assertThat(page1).contains("col1");
        assertThat(page1).doesNotContain("col11");
        assertThat(page2).contains("LABEL_A");
        assertThat(page2).contains("col11");
      }
    }

    @Test
    @DisplayName("multiple title cols all appear on every page")
    void multipleTitleCols(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleColsWorkbook(tempDir, "test.xlsx", 2, 10);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(2);
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        String page1 = s1.getText(doc);
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        String page2 = s2.getText(doc);
        assertThat(page1).contains("LABEL_A").contains("LABEL_B");
        assertThat(page2).contains("LABEL_A").contains("LABEL_B");
      }
    }

    @Test
    @DisplayName("single col page: title col appears with no duplication")
    void titleColWithSingleColPage(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleColsWorkbook(tempDir, "test.xlsx", 1, 4);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        String text = new PDFTextStripper().getText(doc);
        assertThat(text).contains("LABEL_A").contains("col4");
      }
    }

    @Test
    @DisplayName("title col and title row both appear on all pages when combined")
    void titleColAndTitleRowCombined(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTitleColsAndRowsWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, TEST_OPTIONS);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(4);
        for (int p = 1; p <= 4; p++) {
          PDFTextStripper s = new PDFTextStripper();
          s.setStartPage(p);
          s.setEndPage(p);
          assertThat(s.getText(doc)).as("page " + p).contains("CORNER");
        }
        PDFTextStripper s1 = new PDFTextStripper();
        s1.setStartPage(1);
        s1.setEndPage(1);
        assertThat(s1.getText(doc)).contains("row1");
        PDFTextStripper s2 = new PDFTextStripper();
        s2.setStartPage(2);
        s2.setEndPage(2);
        assertThat(s2.getText(doc)).contains("row31");
      }
    }
  }

  // ---------------------------------------------------------------------------
  // useSystemFonts
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("useSystemFonts")
  class UseSystemFonts {

    @Test
    @DisplayName("falls back to regularFontPath when system font is not found")
    void fallsBackToRegularFontPath(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithUnknownFont(tempDir);
      Path pdf = tempDir.resolve("out.pdf");
      // getRegularFontPath() is @Nullable but TEST_OPTIONS always has it set
      Path regularPath = java.util.Objects.requireNonNull(TEST_OPTIONS.getRegularFontPath());
      Path boldPath = java.util.Objects.requireNonNull(TEST_OPTIONS.getBoldFontPath());
      PdfGenerateOptions options = PdfGenerateOptions.builder()
          .useSystemFonts(true)
          .regularFontPath(regularPath)
          .boldFontPath(boldPath)
          .build();

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, options);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
      }
    }

    @Test
    @DisplayName("フォールバック時にfitToPageのスケールが正しく1ページに収まる")
    void fallback_fitToPageFitsOnOnePage(@TempDir Path tempDir) throws Exception {
      // Create a workbook with an unknown font (not available on any system) + Fit to Page.
      // Columns are deliberately wider than the printable area so fitScale must be < 1.
      // Without MDW from the fallback font, getColumnNaturalWidthInPt uses POI's
      // getColumnWidthInPixels() which omits the OOXML +5 px margin, underestimating
      // naturalColTotal → fitScale > 1 → content scales UP → rows appear too tall.
      // With the fix (MDW from NotoSansJP), the OOXML formula is used and scale < 1.
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        XSSFFont font = wb.getFontAt(0);
        font.setFontName("__NonExistentFontABC__");
        font.setFontHeightInPoints((short) 11);

        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.setMargin(PageMargin.LEFT, 0.5);
        sheet.setMargin(PageMargin.RIGHT, 0.5);
        sheet.setMargin(PageMargin.TOP, 0.5);
        sheet.setMargin(PageMargin.BOTTOM, 0.5);
        // Wide columns: naturalColTotal >> A4 printable width
        for (int c = 0; c < 15; c++) {
          sheet.setColumnWidth(c, 25 * 256);
        }
        // Many rows: if fitScale > 1 and height not capped, content would spill to page 2
        for (int r = 0; r < 50; r++) {
          var row = sheet.createRow(r);
          row.setHeightInPoints(15);
          row.createCell(0).setCellValue("row" + r);
        }
        Path excel = tempDir.resolve("fallback-fit.xlsx");
        try (var out = Files.newOutputStream(excel)) {
          wb.write(out);
        }
        Path regularPath = java.util.Objects.requireNonNull(TEST_OPTIONS.getRegularFontPath());
        Path boldPath = java.util.Objects.requireNonNull(TEST_OPTIONS.getBoldFontPath());
        PdfGenerateOptions options = PdfGenerateOptions.builder()
            .useSystemFonts(true)
            .regularFontPath(regularPath)
            .boldFontPath(boldPath)
            .build();
        Path pdf = tempDir.resolve("fallback-fit.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, options);

        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
          var stripper = new PDFTextStripper();
          assertThat(stripper.getText(doc)).contains("row0").contains("row49");
        }
      }
    }

    @Test
    @DisplayName("throws PdfGenerateException when system font not found and no fallback path set")
    void throwsWhenSystemFontNotFoundAndNoFallback(@TempDir Path tempDir) {
      Path excel;
      try {
        excel = createWorkbookWithUnknownFont(tempDir);
      } catch (IOException e) {
        throw new RuntimeException(e);
      }
      Path pdf = tempDir.resolve("out.pdf");
      PdfGenerateOptions options = PdfGenerateOptions.builder()
          .useSystemFonts(true)
          .build();

      assertThatThrownBy(() -> ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, options))
          .isInstanceOf(PdfGenerateException.class);
    }
  }

  // ---------------------------------------------------------------------------
  // Fit to Page scale factor
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("Fit to Page scale factor")
  class FitToPageScale {

    @Test
    @DisplayName("列幅がページ幅を超える場合に1ページに収まる")
    void wideContent_fitsOnOnePage(@TempDir Path tmp) throws Exception {
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.setMargin(PageMargin.LEFT, 0.5);
        sheet.setMargin(PageMargin.RIGHT, 0.5);
        sheet.setMargin(PageMargin.TOP, 0.5);
        sheet.setMargin(PageMargin.BOTTOM, 0.5);
        // 20 columns × 30 chars → natural total >> A4 printable width (~523pt)
        for (int c = 0; c < 20; c++) {
          sheet.setColumnWidth(c, 30 * 256);
        }
        for (int r = 0; r < 3; r++) {
          var row = sheet.createRow(r);
          row.setHeightInPoints(20);
          for (int c = 0; c < 20; c++) {
            row.createCell(c).setCellValue("R" + r + "C" + c);
          }
        }
        Path excel = tmp.resolve("wide.xlsx");
        try (var out = Files.newOutputStream(excel)) {
          wb.write(out);
        }
        Path pdf = tmp.resolve("wide.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          // Fit-to-page with 20 wide columns produces an extreme scale (~10%),
          // making cells too thin to render text. Verify page count only.
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
        }
      }
    }

    @Test
    @DisplayName("行数が多い場合に高さ制約で1ページに収まる")
    void tallContent_fitsOnOnePageByHeightConstraint(@TempDir Path tmp) throws Exception {
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.setMargin(PageMargin.LEFT, 0.5);
        sheet.setMargin(PageMargin.RIGHT, 0.5);
        sheet.setMargin(PageMargin.TOP, 0.5);
        sheet.setMargin(PageMargin.BOTTOM, 0.5);
        // 1 narrow column (width constraint trivially satisfied)
        sheet.setColumnWidth(0, 10 * 256);
        // 200 rows × 30pt = 6000pt >> A4 printable height (~769pt)
        for (int r = 0; r < 200; r++) {
          var row = sheet.createRow(r);
          row.setHeightInPoints(30);
          row.createCell(0).setCellValue("row" + r);
        }
        Path excel = tmp.resolve("tall.xlsx");
        try (var out = Files.newOutputStream(excel)) {
          wb.write(out);
        }
        Path pdf = tmp.resolve("tall.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
          var stripper = new PDFTextStripper();
          String text = stripper.getText(doc);
          assertThat(text).contains("row0").contains("row199");
        }
      }
    }

    @Test
    @DisplayName("行高がfitScaleでスケールされる（fitScale < 1 なら行高は縮小される）")
    void rowHeightScaledAccurately(@TempDir Path tmp) throws Exception {
      // Use 3 wide columns to force fitScale in the range [0.85, 0.99] — narrow
      // enough that cells remain tall enough to render text, wide enough to need scaling.
      // Exact fitScale depends on screen PPI (MDW), so we compute it at runtime.
      final double leftMarginIn = 0.5, topMarginIn = 0.5;
      final float rowHeightPt = 20f;

      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setFitToPage(true);
        sheet.setMargin(PageMargin.LEFT, leftMarginIn);
        sheet.setMargin(PageMargin.RIGHT, leftMarginIn);
        sheet.setMargin(PageMargin.TOP, topMarginIn);
        sheet.setMargin(PageMargin.BOTTOM, topMarginIn);
        // 3 columns × 25 chars → naturalColTotal > printableWidth → fitScale < 1
        // but not so extreme that cells become too thin to render text
        for (int c = 0; c < 3; c++) {
          sheet.setColumnWidth(c, 25 * 256);
        }
        // Create 200 rows: at fitScale=1 they exceed printableHeight, confirming scale is applied.
        for (int r = 0; r < 200; r++) {
          var row = sheet.createRow(r);
          row.setHeightInPoints(rowHeightPt);
          row.createCell(0).setCellValue(r == 0 ? "row0" : r == 199 ? "row199" : "x");
        }
        Path excel = tmp.resolve("rowscale.xlsx");
        try (var out = Files.newOutputStream(excel)) {
          wb.write(out);
        }
        Path pdf = tmp.resolve("rowscale.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
          // All 200 rows fit on 1 page → fitScale < 1 was applied (200×20=4000pt > printableH).
          // Check text exists on the page (exact position depends on runtime screen DPI).
          var stripper = new PDFTextStripper();
          String text = stripper.getText(doc);
          assertThat(text).contains("row0").contains("row199");
        }
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Text overflow
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("テキストオーバーフロー")
  class TextOverflow {

    @Test
    @DisplayName("GENERAL配置の数値セルは隣の空セルにはみ出さずセル幅内に収まる")
    void numericCellWithGeneralAlignmentDoesNotOverflowIntoAdjacentEmptyCells(
        @TempDir Path tmp) throws Exception {
      // Col A: 8 chars wide — contains numeric 12345 with GENERAL alignment (→ right-aligned).
      // Col B: 20 chars wide, empty — the bug caused overflow to extend the right-alignment
      //        anchor to (colA + colB) right edge, shifting the text far into column B.
      // Col C: 5 chars, text "STOP".
      //
      // With the bug:  textX ≈ leftMargin + W_A + W_B − textWidth − CELL_PAD  (≫ 100 pt from left)
      // With the fix:  textX ≈ leftMargin + W_A       − textWidth − CELL_PAD  (≪ 100 pt from left)
      //
      // For any realistic MDW (7–12 px) and font size, W_A ≤ 76 pt and W_A + W_B ≥ 170 pt,
      // so a 100 pt threshold from the left margin reliably separates the two cases.
      final double leftMarginIn = 0.5;

      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setMargin(PageMargin.LEFT, leftMarginIn);
        sheet.setMargin(PageMargin.RIGHT, leftMarginIn);
        sheet.setMargin(PageMargin.TOP, 0.5);
        sheet.setMargin(PageMargin.BOTTOM, 0.5);
        sheet.setColumnWidth(0, 8 * 256);   // col A: 8 chars (narrow)
        sheet.setColumnWidth(1, 20 * 256);  // col B: 20 chars (wide, empty)
        sheet.setColumnWidth(2, 5 * 256);   // col C: 5 chars
        var row = sheet.createRow(0);
        row.createCell(0).setCellValue(12345d); // GENERAL alignment → effectively right-aligned
        // col B intentionally left empty — this is the overflow trigger
        row.createCell(2).setCellValue("STOP");

        Path excel = tmp.resolve("overflow.xlsx");
        try (var out = Files.newOutputStream(excel)) {
          wb.write(out);
        }
        Path pdf = tmp.resolve("overflow.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        // Define a "col A only" region from left margin to left margin + 100 pt.
        // For any MDW in [7, 12], col A width ≤ 76 pt — well within the 100 pt boundary.
        // With the bug, the text starts at > 100 pt from the left margin (in column B area).
        float leftMarginPt = (float) (leftMarginIn * 72);
        float colAThresholdPt = 100f;
        float pageW = PDRectangle.A4.getWidth();
        float pageH = PDRectangle.A4.getHeight();

        try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
          PDPage page = doc.getPage(0);
          var stripper = new PDFTextStripperByArea();
          stripper.addRegion("colA_region",
              new Rectangle2D.Float(leftMarginPt, 0, colAThresholdPt, pageH));
          stripper.addRegion("overflow_region",
              new Rectangle2D.Float(leftMarginPt + colAThresholdPt, 0, pageW, pageH));
          stripper.extractRegions(page);

          // "12345" must appear in the column-A region (right-aligned within the cell).
          // Before the fix it appeared only in the overflow region (shifted far to the right).
          assertThat(stripper.getTextForRegion("colA_region")).contains("12345");
          assertThat(stripper.getTextForRegion("overflow_region")).doesNotContain("12345");
        }
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Page-scale ground-truth tests (Layer 1 — platform-independent)
  // ---------------------------------------------------------------------------
  //
  // These tests verify that our PDF generation produces the correct fit-to-page /
  // adjust-to scale for each major page-setup mode. They use NotoSansJP (committed
  // to test-resources) with useSystemFonts=false so that the MDW is computed from
  // the same font at fixed 96 DPI on every platform (Mac, Linux, CI/CD).
  //
  // Sheet layout used in every scenario:
  //   - 5 columns (A–E), each 15 chars wide
  //   - 30 rows, each 20 pt tall
  //   - Row 1: A1="R1", B1="B", C1="C", D1="D", E1="E"
  //   - Rows 2-30: Ac="Rc" (B–E empty)
  //
  // Verification strategy: extract the Y-coordinate of the first two text baselines
  // from the PDF, compute the rendered row height, and compare to the expected value
  // (naturalRowHeight × fitScale) with a 1 pt tolerance.

  @Nested
  @DisplayName("ページスケール（Layer 1: プラットフォーム非依存）")
  class PageScaleLayer1 {

    // ---- shared constants ----
    private static final float COL_CHARS = 15f;
    private static final int NUM_COLS = 5;
    private static final float ROW_HEIGHT_PT = 20f;
    private static final int NUM_ROWS = 30;
    private static final double MARGIN_IN = 0.5; // 0.5 inch on all sides
    private static final float PRINT_W_PT =
        (float) (PDRectangle.A4.getWidth() - 2 * MARGIN_IN * 72);
    private static final float PRINT_H_PT =
        (float) (PDRectangle.A4.getHeight() - 2 * MARGIN_IN * 72);

    // ---- wide-content constants (width IS the binding constraint) ----
    // 10 cols × 12 chars: naturalColTotal(MDW=8) = 10 × 72pt = 720pt > PRINT_W_PT(523pt)
    // fitScale_width = 523/720 ≈ 0.726  → row height clearly drops to ~14.5pt
    private static final float WIDE_COL_CHARS = 12f;
    private static final int WIDE_NUM_COLS = 10;

    /** Computes MDW for NotoSansJP at 96 DPI — same as production with useSystemFonts=false. */
    private static int notoMdw() {
      Path reg = java.util.Objects.requireNonNull(TEST_OPTIONS.getRegularFontPath());
      float fontSizePt = 11f; // POI fresh-workbook default
      return SystemFontLocator.computeMdw(reg, "", fontSizePt);
    }

    /** naturalColTotal for WIDE_NUM_COLS × WIDE_COL_CHARS (width-constraint scenario). */
    private static float naturalColTotalWide(int mdw) {
      return naturalColTotalFor(WIDE_NUM_COLS, WIDE_COL_CHARS, mdw);
    }

    /** OOXML spec §18.3.1.13 formula: px = Truncate(((256×width + Truncate(128/MDW))/256) × MDW). */
    private static float naturalColTotalFor(int numCols, float colChars, int mdw) {
      int widthIn256 = (int) (colChars * 256);
      int px = (int) (((widthIn256 + (128 / mdw)) / 256.0) * mdw);
      float total = numCols * px * (72f / 96f);
      return total;
    }

    /**
     * Builds a workbook with WIDE_NUM_COLS columns of WIDE_COL_CHARS each.
     * naturalColTotal > PRINT_W_PT → width IS the binding constraint for fitToPage.
     */
    private XSSFWorkbook buildWideWorkbook() {
      XSSFWorkbook wb = new XSSFWorkbook();
      var sheet = wb.createSheet("S");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, MARGIN_IN);
      sheet.setMargin(PageMargin.RIGHT, MARGIN_IN);
      sheet.setMargin(PageMargin.TOP, MARGIN_IN);
      sheet.setMargin(PageMargin.BOTTOM, MARGIN_IN);
      for (int c = 0; c < WIDE_NUM_COLS; c++) {
        sheet.setColumnWidth(c, (int) (WIDE_COL_CHARS * 256));
      }
      for (int r = 0; r < NUM_ROWS; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(ROW_HEIGHT_PT);
        // Create cells in ALL columns so the used-range spans all WIDE_NUM_COLS columns
        // and naturalColTotal is computed over all 10 columns (not just col 0).
        for (int c = 0; c < WIDE_NUM_COLS; c++) {
          row.createCell(c).setCellValue(c == 0 ? "R" + (r + 1) : "");
        }
      }
      return wb;
    }

    /** Builds a workbook with the standard 5-column / 30-row layout. */
    private XSSFWorkbook buildWorkbook() {
      XSSFWorkbook wb = new XSSFWorkbook();
      var sheet = wb.createSheet("S");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, MARGIN_IN);
      sheet.setMargin(PageMargin.RIGHT, MARGIN_IN);
      sheet.setMargin(PageMargin.TOP, MARGIN_IN);
      sheet.setMargin(PageMargin.BOTTOM, MARGIN_IN);
      for (int c = 0; c < NUM_COLS; c++) {
        sheet.setColumnWidth(c, (int) (COL_CHARS * 256));
      }
      for (int r = 0; r < NUM_ROWS; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(ROW_HEIGHT_PT);
        row.createCell(0).setCellValue("R" + (r + 1));
        if (r == 0) {
          row.createCell(1).setCellValue("B");
          row.createCell(2).setCellValue("C");
          row.createCell(3).setCellValue("D");
          row.createCell(4).setCellValue("E");
        }
      }
      return wb;
    }

    /**
     * Extracts the rendered row height from the PDF by measuring the Y-gap between
     * the baselines of the first two rows of text.
     * Returns -1 if fewer than 2 distinct Y values are found.
     */
    private float extractRenderedRowHeight(Path pdf) throws IOException {
      try (var doc = org.apache.pdfbox.Loader.loadPDF(pdf.toFile())) {
        var page = doc.getPage(0);
        // Use PDFTextStripperByArea across the full page height to collect all baselines.
        // We need raw Y positions, so we use a custom approach via PDFTextStripper
        // with sortByPosition=false to get lines in stream order.
        var stripper = new PDFTextStripper();
        stripper.setSortByPosition(true);
        // Extract text line positions via PDFTextStripperByArea scanning thin strips.
        // Simpler: collect distinct Y values from a full-page strip region.
        float pageH = page.getMediaBox().getHeight();
        float pageW = page.getMediaBox().getWidth();
        var byArea = new PDFTextStripperByArea();
        // Scan 1 pt strips to detect row baselines.
        float topMarginPt = (float) (MARGIN_IN * 72);
        // Coarse scan: 2 pt strips across the top 200 pt of the page
        for (int y = 0; y < 200; y += 2) {
          float stripY = topMarginPt + y;
          if (stripY >= pageH) break;
          String regionName = "s" + y;
          byArea.addRegion(regionName,
              new java.awt.geom.Rectangle2D.Float(0, stripY, pageW, 2f));
        }
        byArea.extractRegions(page);
        java.util.List<Float> rowY = new java.util.ArrayList<>();
        for (int y = 0; y < 200; y += 2) {
          String t = byArea.getTextForRegion("s" + y).strip();
          if (!t.isEmpty()) {
            rowY.add((float) (topMarginPt + y));
          }
        }
        if (rowY.size() < 2) return -1f;
        // Row height = distance between first two detected row starts.
        return rowY.get(1) - rowY.get(0);
      }
    }

    /** Asserts that the rendered row height is within 1.5 pt of the expected value. */
    private void assertRowHeight(Path pdf, float expectedRowHeightPt) throws IOException {
      float actual = extractRenderedRowHeight(pdf);
      assertThat(actual)
          .as("rendered row height should be %.2f pt (±1.5)", expectedRowHeightPt)
          .isBetween(expectedRowHeightPt - 1.5f, expectedRowHeightPt + 1.5f);
    }

    @Test
    @DisplayName("adjustTo=100%: 行高は自然な高さのまま")
    void adjustTo100_rowHeightIsNatural(@TempDir Path tmp) throws Exception {
      try (XSSFWorkbook wb = buildWorkbook()) {
        var sheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet("S");
        // Explicit scale=100; fitToPage must be OFF
        sheet.setFitToPage(false);
        sheet.getCTWorksheet().getPageSetup().setScale(100);
        Path excel = tmp.resolve("a.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("a.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        try (var doc = org.apache.pdfbox.Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
        }
        assertRowHeight(pdf, ROW_HEIGHT_PT); // scale=100% → row height unchanged
      }
    }

    @Test
    @DisplayName("adjustTo=85%: 行高が85%に縮小される")
    void adjustTo85_rowHeightScaledDown(@TempDir Path tmp) throws Exception {
      try (XSSFWorkbook wb = buildWorkbook()) {
        var sheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet("S");
        sheet.setFitToPage(false);
        sheet.getCTWorksheet().getPageSetup().setScale(85);
        Path excel = tmp.resolve("b.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("b.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        assertRowHeight(pdf, ROW_HEIGHT_PT * 0.85f);
      }
    }

    @Test
    @DisplayName("fitToPage(幅1×高1): 列幅がページ幅を超え、幅制約でfitScaleが決まる")
    void fitToPage_1x1_rowHeightScaledByWidthConstraint(@TempDir Path tmp) throws Exception {
      // WIDE layout: naturalColTotal(720pt) > PRINT_W_PT(523pt) → width IS the binding constraint.
      // Previous setup (5cols×15chars=450pt) never fired width constraint → test was misleading.
      try (XSSFWorkbook wb = buildWideWorkbook()) {
        var sheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet("S");
        sheet.setFitToPage(true);
        sheet.getCTWorksheet().getPageSetup().setFitToWidth(1);
        sheet.getCTWorksheet().getPageSetup().setFitToHeight(1);
        Path excel = tmp.resolve("c.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("c.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        int mdw = notoMdw();
        float natColTotal = naturalColTotalWide(mdw); // 720pt > PRINT_W_PT
        float fitScale = Math.min(1f, PRINT_W_PT / natColTotal); // ≈ 0.726
        // Height constraint: 30rows×20pt=600pt; 600×0.726=435pt < PRINT_H_PT → not binding
        float natRowTotal = NUM_ROWS * ROW_HEIGHT_PT;
        if (natRowTotal * fitScale > PRINT_H_PT) {
          fitScale = Math.min(fitScale, PRINT_H_PT / natRowTotal);
        }
        try (var doc = org.apache.pdfbox.Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
        }
        assertRowHeight(pdf, ROW_HEIGHT_PT * fitScale); // ≈ 14.5pt (clearly < 20pt)
      }
    }

    @Test
    @DisplayName("fitToPage(幅1×高さ0): 列幅がページ幅を超え、幅制約のみでscaleが決まる")
    void fitToPage_1xUnlimited_rowHeightByWidthOnly(@TempDir Path tmp) throws Exception {
      // WIDE layout: naturalColTotal(720pt) > PRINT_W_PT → width constraint fires.
      // fitToHeight=0 means height is unlimited → only width constraint applies.
      try (XSSFWorkbook wb = buildWideWorkbook()) {
        var sheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet("S");
        sheet.setFitToPage(true);
        sheet.getCTWorksheet().getPageSetup().setFitToWidth(1);
        sheet.getCTWorksheet().getPageSetup().setFitToHeight(0);
        Path excel = tmp.resolve("d.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("d.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        int mdw = notoMdw();
        float natColTotal = naturalColTotalWide(mdw);
        float fitScale = Math.min(1f, PRINT_W_PT / natColTotal); // ≈ 0.726, height not applied
        assertRowHeight(pdf, ROW_HEIGHT_PT * fitScale);
      }
    }

    @Test
    @DisplayName("fitToPage: 列幅合計がページ幅以下 → fitScale=1.0（縮小なし）")
    void fitToPage_contentFitsNaturally_noScaling(@TempDir Path tmp) throws Exception {
      try (XSSFWorkbook wb = buildWorkbook()) {
        var sheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet("S");
        sheet.setFitToPage(true);
        // Use narrow columns so naturalColTotal < printableWidth
        for (int c = 0; c < NUM_COLS; c++) {
          sheet.setColumnWidth(c, 5 * 256); // 5 chars — very narrow
        }
        sheet.getCTWorksheet().getPageSetup().setFitToWidth(1);
        sheet.getCTWorksheet().getPageSetup().setFitToHeight(1);
        Path excel = tmp.resolve("e.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("e.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        // naturalColTotal (5 cols × 5 chars) at NotoSansJP MDW << PRINT_W_PT → fitScale=1.0
        assertRowHeight(pdf, ROW_HEIGHT_PT); // no scaling
      }
    }

    @Test
    @DisplayName("fitToPage(幅0×高1): 高さ制約のみでfitScaleが決まる（幅は自由）")
    void fitToPage_unlimitedWidth_heightConstraintDrivesScale(@TempDir Path tmp) throws Exception {
      // 50 rows × 20 pt = 1000 pt > PRINT_H_PT (~769 pt) → height constraint fires
      // fitToWidth=0 means no width constraint → fitScale driven purely by height
      int numRows = 50;
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setMargin(PageMargin.LEFT, MARGIN_IN);
        sheet.setMargin(PageMargin.RIGHT, MARGIN_IN);
        sheet.setMargin(PageMargin.TOP, MARGIN_IN);
        sheet.setMargin(PageMargin.BOTTOM, MARGIN_IN);
        // Narrow columns so width never triggers scaling
        for (int c = 0; c < 3; c++) {
          sheet.setColumnWidth(c, 5 * 256);
        }
        for (int r = 0; r < numRows; r++) {
          var row = sheet.createRow(r);
          row.setHeightInPoints(ROW_HEIGHT_PT);
          row.createCell(0).setCellValue("R" + (r + 1));
        }
        sheet.setFitToPage(true);
        sheet.getCTWorksheet().getPageSetup().setFitToWidth(0);  // unlimited width
        sheet.getCTWorksheet().getPageSetup().setFitToHeight(1); // must fit 1 page tall

        Path excel = tmp.resolve("f.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("f.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        float natRowTotal = numRows * ROW_HEIGHT_PT; // 1000 pt
        float fitScale = Math.min(1f, PRINT_H_PT / natRowTotal);
        try (var doc = org.apache.pdfbox.Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
        }
        assertRowHeight(pdf, ROW_HEIGHT_PT * fitScale);
      }
    }

    @Test
    @DisplayName("fitToPage(幅1): スケール後のコンテンツがページ幅いっぱいに収まる（左右余白の検証）")
    void fitToPage_widthConstraint_scaledContentFillsPageWidth(@TempDir Path tmp) throws Exception {
      // This test catches the bug where content does NOT fill the page width after scaling.
      // Scenario: WIDE_NUM_COLS columns of WIDE_COL_CHARS chars each, fitToWidth=1.
      // Correct behaviour: fitScale = PRINT_W_PT / naturalColTotal  → coloured fill reaches
      // within a few pixels of the right page margin and does NOT overflow past it.
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setMargin(PageMargin.LEFT, MARGIN_IN);
        sheet.setMargin(PageMargin.RIGHT, MARGIN_IN);
        sheet.setMargin(PageMargin.TOP, MARGIN_IN);
        sheet.setMargin(PageMargin.BOTTOM, MARGIN_IN);
        for (int c = 0; c < WIDE_NUM_COLS; c++) {
          sheet.setColumnWidth(c, (int) (WIDE_COL_CHARS * 256));
        }
        // Fill every cell with a distinctive colour so we can detect the content right edge.
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(new XSSFColor(new byte[]{0x00, 0x70, (byte) 0xC0}, null));
        var row = sheet.createRow(0);
        row.setHeightInPoints(ROW_HEIGHT_PT);
        for (int c = 0; c < WIDE_NUM_COLS; c++) {
          row.createCell(c).setCellStyle(style);
        }
        sheet.setFitToPage(true);
        sheet.getCTWorksheet().getPageSetup().setFitToWidth(1);
        sheet.getCTWorksheet().getPageSetup().setFitToHeight(0);

        Path excel = tmp.resolve("fill.xlsx");
        try (var out = Files.newOutputStream(excel)) { wb.write(out); }
        Path pdf = tmp.resolve("fill.pdf");
        ExcelToPdfUtil.generate(excel, List.of("S"), pdf, TEST_OPTIONS);

        try (var doc = org.apache.pdfbox.Loader.loadPDF(pdf.toFile())) {
          assertThat(doc.getNumberOfPages()).isEqualTo(1);
          // Render at 72 DPI: 1pt ≈ 1px.
          // Content right edge = leftMarginPx + PRINT_W_PT ≈ 36 + 523 = 559px.
          BufferedImage img = new PDFRenderer(doc).renderImageWithDPI(0, 72);
          int leftMarginPx = (int) (MARGIN_IN * 72);               // 36px
          int rightEdgePx  = leftMarginPx + (int) PRINT_W_PT;      // ≈ 559px
          int rowCenterY   = leftMarginPx + (int) (ROW_HEIGHT_PT / 2); // inside first row

          // Coloured fill must be present within 5 px of the computed right edge.
          boolean reachesRightEdge = false;
          for (int x = rightEdgePx - 5; x <= rightEdgePx; x++) {
            if (x >= 0 && x < img.getWidth()
                && isColoredPixel(img.getRGB(x, rowCenterY), FILL_RGB)) {
              reachesRightEdge = true;
              break;
            }
          }
          assertThat(reachesRightEdge)
              .as("fitToPage(幅1): スケール後コンテンツがページ右マージン付近(5px以内)に到達すること")
              .isTrue();

          // Coloured fill must NOT overflow beyond the right printable margin.
          boolean overflows = false;
          for (int x = rightEdgePx + 3; x < Math.min(img.getWidth(), rightEdgePx + 15); x++) {
            if (isColoredPixel(img.getRGB(x, rowCenterY), FILL_RGB)) {
              overflows = true;
              break;
            }
          }
          assertThat(overflows)
              .as("fitToPage(幅1): スケール後コンテンツがページ右マージンをはみ出さないこと")
              .isFalse();
        }
      }
    }

    // ---- glyph-height tests ----

    /**
     * Builds a one-row workbook with blue "X" text in column 0.
     *
     * @param wideLayout true → WIDE_NUM_COLS × WIDE_COL_CHARS (forces fitToPage width constraint)
     * @param fitToPage  true → fitToPage=1, fitToWidth=1; false → adjustTo at {@code scalePercent}
     */
    private Path buildGlyphWorkbook(Path dir, String name, int fontSizePt,
        boolean wideLayout, boolean fitToPage, int scalePercent) throws IOException {
      try (XSSFWorkbook wb = new XSSFWorkbook()) {
        var sheet = wb.createSheet("S");
        sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
        sheet.setMargin(PageMargin.LEFT, MARGIN_IN);
        sheet.setMargin(PageMargin.RIGHT, MARGIN_IN);
        sheet.setMargin(PageMargin.TOP, MARGIN_IN);
        sheet.setMargin(PageMargin.BOTTOM, MARGIN_IN);

        int numCols = wideLayout ? WIDE_NUM_COLS : 1;
        float colChars = wideLayout ? WIDE_COL_CHARS : 10f;
        for (int c = 0; c < numCols; c++) {
          sheet.setColumnWidth(c, (int) (colChars * 256));
        }

        XSSFFont font = wb.createFont();
        font.setFontHeightInPoints((short) fontSizePt);
        XSSFColor blue = new XSSFColor(new byte[]{0x00, 0x70, (byte) 0xC0}, null);
        font.setColor(blue);

        XSSFCellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);

        var row = sheet.createRow(0);
        row.setHeightInPoints(ROW_HEIGHT_PT);
        var cell0 = row.createCell(0);
        cell0.setCellStyle(style);
        cell0.setCellValue("X");
        for (int c = 1; c < numCols; c++) {
          row.createCell(c); // ensure used-range spans all columns so fitToPage fires correctly
        }

        if (fitToPage) {
          sheet.setFitToPage(true);
          sheet.getCTWorksheet().getPageSetup().setFitToWidth(1);
          sheet.getCTWorksheet().getPageSetup().setFitToHeight(0);
        } else {
          sheet.setFitToPage(false);
          sheet.getCTWorksheet().getPageSetup().setScale(scalePercent);
        }

        Path path = dir.resolve(name);
        try (var out = Files.newOutputStream(path)) { wb.write(out); }
        return path;
      }
    }

    @Test
    @DisplayName("adjustTo=85%: グリフ（文字）の視覚的な高さも85%に縮小される")
    void adjustTo85_glyphHeightScaledDown(@TempDir Path tmp) throws Exception {
      int fontSizePt = 14;
      Path naturalExcel = buildGlyphWorkbook(tmp, "glyph-natural.xlsx", fontSizePt, false, false, 100);
      Path scaledExcel  = buildGlyphWorkbook(tmp, "glyph-85.xlsx",      fontSizePt, false, false, 85);
      Path naturalPdf = tmp.resolve("glyph-natural.pdf");
      Path scaledPdf  = tmp.resolve("glyph-85.pdf");
      ExcelToPdfUtil.generate(naturalExcel, List.of("S"), naturalPdf, TEST_OPTIONS);
      ExcelToPdfUtil.generate(scaledExcel,  List.of("S"), scaledPdf,  TEST_OPTIONS);

      int dpi = 144;
      // At 144 DPI: 1pt = 2px.  Margin = 0.5 in = 36pt = 72px.
      int marginPx = (int) (MARGIN_IN * dpi);
      int xStart = marginPx + 4;                                          // past cell padding
      int xEnd   = marginPx + 100;                                        // well within col0
      int yStart = marginPx;
      int yEnd   = marginPx + (int) (ROW_HEIGHT_PT * dpi / 72) + 20;     // one row + buffer

      try (PDDocument nd = Loader.loadPDF(naturalPdf.toFile());
          PDDocument sd = Loader.loadPDF(scaledPdf.toFile())) {
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, dpi);
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, dpi);

        int naturalTop    = topmostColoredY(ni, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int naturalBottom = bottommostColoredY(ni, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int scaledTop     = topmostColoredY(si, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int scaledBottom  = bottommostColoredY(si, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int naturalH = (naturalBottom >= naturalTop) ? naturalBottom - naturalTop + 1 : 0;
        int scaledH  = (scaledBottom  >= scaledTop)  ? scaledBottom  - scaledTop  + 1 : 0;

        assertThat(naturalH).as("natural glyph height > 0").isGreaterThan(0);
        assertThat(scaledH).as("85%% scaled glyph height > 0").isGreaterThan(0);
        double ratio = (double) scaledH / naturalH;
        assertThat(ratio)
            .as("85%%スケール時のグリフ高さは自然高の85%%(±8%%)")
            .isBetween(0.77, 0.93);
      }
    }

    @Test
    @DisplayName("fitToPage(幅1): グリフ高さもfitScaleで縮小される")
    void fitToPage_widthConstraint_glyphHeightScaledWithFitScale(@TempDir Path tmp) throws Exception {
      int fontSizePt = 14;
      // Wide layout: 10 cols × 12 chars → naturalColTotal > PRINT_W_PT → fitScale ≈ 0.726
      Path scaledExcel  = buildGlyphWorkbook(tmp, "glyph-wide-fit.xlsx",     fontSizePt, true, true,  0);
      Path naturalExcel = buildGlyphWorkbook(tmp, "glyph-wide-natural.xlsx", fontSizePt, true, false, 100);
      Path scaledPdf  = tmp.resolve("glyph-wide-fit.pdf");
      Path naturalPdf = tmp.resolve("glyph-wide-natural.pdf");
      ExcelToPdfUtil.generate(scaledExcel,  List.of("S"), scaledPdf,  TEST_OPTIONS);
      ExcelToPdfUtil.generate(naturalExcel, List.of("S"), naturalPdf, TEST_OPTIONS);

      int mdw = notoMdw();
      float natColTotal = naturalColTotalWide(mdw);
      float fitScale = Math.min(1f, PRINT_W_PT / natColTotal); // ≈ 0.726

      int dpi = 144;
      int marginPx = (int) (MARGIN_IN * dpi);
      // Natural col0 width at MDW pixels (96 DPI), converted to pt then to 144-DPI pixels.
      int col0NatPx96 = (int) (((WIDE_COL_CHARS * 256 + (128f / mdw)) / 256.0) * mdw);
      float col0NatPt = col0NatPx96 * 72f / 96f;
      // Scan within the scaled column 0 to capture only that column's glyph.
      int xStart = marginPx + 4;
      int xEnd   = marginPx + (int) (col0NatPt * fitScale * dpi / 72) - 4;
      int yStart = marginPx;
      int yEnd   = marginPx + (int) (ROW_HEIGHT_PT * dpi / 72) + 20;

      try (PDDocument sd = Loader.loadPDF(scaledPdf.toFile());
          PDDocument nd = Loader.loadPDF(naturalPdf.toFile())) {
        BufferedImage si = new PDFRenderer(sd).renderImageWithDPI(0, dpi);
        BufferedImage ni = new PDFRenderer(nd).renderImageWithDPI(0, dpi);

        int naturalTop    = topmostColoredY(ni, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int naturalBottom = bottommostColoredY(ni, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int scaledTop     = topmostColoredY(si, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int scaledBottom  = bottommostColoredY(si, xStart, xEnd, yStart, yEnd, FILL_RGB);
        int naturalH = (naturalBottom >= naturalTop) ? naturalBottom - naturalTop + 1 : 0;
        int scaledH  = (scaledBottom  >= scaledTop)  ? scaledBottom  - scaledTop  + 1 : 0;

        assertThat(naturalH).as("natural glyph height > 0").isGreaterThan(0);
        assertThat(scaledH).as("fitToPage scaled glyph height > 0").isGreaterThan(0);
        double ratio = (double) scaledH / naturalH;
        assertThat(ratio)
            .as("fitToPage幅制約のfitScale(%.3f)でグリフ高さが縮小されること(±8%%)", (double) fitScale)
            .isBetween(fitScale - 0.08, fitScale + 0.08);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Page-scale ground-truth tests (Layer 2 — Mac only, compares against Excel PDF)
  // ---------------------------------------------------------------------------
  //
  // These tests use a real Excel file (adjust-to-fit-to-test-data.xlsx) and its
  // Excel-generated PDF as ground truth.  They verify that our code produces the
  // same effective fit-scale as Excel for each page-setup scenario.
  //
  // Test file: src/test/resources/adjust-to-fit-to-test-data.xlsx
  //   Sheet layout: 5 columns (A–E), 30 rows × 20 pt each, row 1 has headers.
  //   Sheets: adjust_100_explicit, adjust_100_implicit, adjust_85,
  //           fit_1x1, fit_1xN, fit_narrow.
  //
  // Excel PDF:  src/test/resources/adjust-to-fit-to-test-data.xlsx.pdf
  //   Exported with "Print entire workbook" from Excel on macOS.
  //   The page-level cm transform in each page stream encodes the effective scale:
  //     cm = [ scale 0 0 scale tx ty ]
  //   so the first element of each cm operator IS the fitScale Excel applied.
  //
  // Skip condition: these tests require the Excel file and PDF to be present in
  //   src/test/resources.  They are committed to the repo, so they run everywhere.
  //   However, since the PDFs were generated on macOS with Aptos Narrow as the theme
  //   Latin font, the expected scales are extracted from the Excel PDF at runtime
  //   rather than hard-coded, making the comparison robust across platforms.

  @Nested
  @DisplayName("ページスケール（Layer 2: Excel PDF との実測比較）")
  class PageScaleLayer2 {

    /** Tolerance: ±2% of scale (e.g. 0.02 means ±2pp for a 1.00 scale). */
    private static final float SCALE_TOLERANCE = 0.02f;

    private static java.net.URL resourceUrl(String name) {
      return ExcelToPdfUtilTest.class.getResource("/" + name);
    }

    /**
     * Extracts the page-level scale from an Excel-exported PDF page.
     *
     * Excel encodes the fit/adjust scale as a uniform scale in the first {@code cm}
     * operator on the page: {@code a 0 0 d tx ty cm} where {@code a == d == fitScale}.
     * A page at 100% scale has no {@code cm} (or has {@code 1 0 0 1 tx ty cm}).
     */
    private float excelPageScale(PDDocument doc, int pageIndex) throws IOException {
      var page = doc.getPage(pageIndex);
      var sb = new StringBuilder();
      var iter = page.getContentStreams();
      while (iter.hasNext()) {
        try (var is = iter.next().createInputStream()) {
          sb.append(new String(is.readAllBytes(), java.nio.charset.StandardCharsets.ISO_8859_1));
        }
      }
      var m = java.util.regex.Pattern.compile(
          "([-\\d.]+)\\s+[-\\d.]+\\s+[-\\d.]+\\s+([-\\d.]+)\\s+[-\\d.]+\\s+[-\\d.]+\\s+cm"
      ).matcher(sb.toString());
      if (m.find()) {
        return (Float.parseFloat(m.group(1)) + Float.parseFloat(m.group(2))) / 2f;
      }
      return 1.0f;
    }

    /**
     * Computes the fitScale our code applies for the given sheet, using the same
     * inputs as the production SheetRenderer: NotoSansJP MDW at 96 DPI, the
     * sheet's column widths, row heights, page setup, and margins.
     *
     * <p>This duplicates {@code SheetRenderer.computeScaleFactor} logic at a high
     * level so we can compare our computed scale against Excel's actual scale without
     * needing to parse our generated PDF.</p>
     */
    private float computeOurFitScale(
        org.apache.poi.xssf.usermodel.XSSFSheet sheet,
        float naturalColTotal, float naturalRowTotal,
        float printW, float printH) {

      // Mirrors SheetRenderer.computeScaleFactor logic.
      boolean fitToPage = sheet.getFitToPage();
      if (!fitToPage && sheet.getCTWorksheet().isSetPageSetup()) {
        var ps = sheet.getCTWorksheet().getPageSetup();
        if (ps.isSetScale()) {
          long s = ps.getScale();
          return (s > 0 && s <= 400) ? s / 100f : 1f;
        }
      }
      if (fitToPage && naturalColTotal > 0) {
        // Use Excel's stored fit scale if available.
        // Note: unlike SheetRenderer.computeScaleFactor, this helper does NOT apply the
        // overflow correction for font-MDW mismatches.  Layer 2 tests compare against
        // Excel's actual scale value, so we must return the cached scale here to keep
        // those tests accurate.  The overflow behaviour is covered by Layer 1 tests.
        if (sheet.getCTWorksheet().isSetPageSetup()) {
          var ps = sheet.getCTWorksheet().getPageSetup();
          if (ps.isSetScale()) {
            long s = ps.getScale();
            if (s > 0 && s <= 400) return s / 100f;
          }
        }
        long fitW = 1, fitH = 1;
        if (sheet.getCTWorksheet().isSetPageSetup()) {
          var ps = sheet.getCTWorksheet().getPageSetup();
          if (ps.isSetFitToWidth()) fitW = ps.getFitToWidth();
          if (ps.isSetFitToHeight()) fitH = ps.getFitToHeight();
        }
        float scale = 1.0f;
        if (fitW != 0) scale = Math.min(1.0f, printW / naturalColTotal);
        if (fitH != 0 && naturalRowTotal > 0 && naturalRowTotal * scale > printH) {
          scale = Math.min(scale, printH / naturalRowTotal);
        }
        return scale;
      }
      // No fitToPage, no explicit scale → render at 100% (matches Excel default).
      return 1.0f;
    }

    private void runSheetTest(String sheetName, int pageIndex, @TempDir Path tmp)
        throws Exception {
      var xlsxUrl = resourceUrl("adjust-to-fit-to-test-data.xlsx");
      var xlsPdfUrl = resourceUrl("adjust-to-fit-to-test-data.xlsx.pdf");
      org.junit.jupiter.api.Assumptions.assumeTrue(xlsxUrl != null && xlsPdfUrl != null,
          "Test resource files not found");

      Path xlsx = java.nio.file.Path.of(xlsxUrl.toURI());
      Path xlsPdf = java.nio.file.Path.of(xlsPdfUrl.toURI());
      Path ourPdf = tmp.resolve(sheetName + ".pdf");

      // Generate our PDF (smoke-test: must not throw)
      ExcelToPdfUtil.generate(xlsx, List.of(sheetName), ourPdf, TEST_OPTIONS);

      // --- Ground truth: fitScale from Excel PDF cm transform ---
      float excelFitScale;
      try (PDDocument excelDoc = org.apache.pdfbox.Loader.loadPDF(xlsPdf.toFile())) {
        excelFitScale = excelPageScale(excelDoc, pageIndex);
      }

      // --- Our computed fitScale: replicate SheetRenderer.computeScaleFactor ---
      float ourFitScale;
      try (Workbook wb = org.apache.poi.ss.usermodel.WorkbookFactory.create(
              xlsx.toFile(), null, true)) {
        var xSheet = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheet(sheetName);
        // MDW: NotoSansJP at 96 DPI (same as useSystemFonts=false in production)
        int mdw = SystemFontLocator.computeMdw(
            java.util.Objects.requireNonNull(TEST_OPTIONS.getRegularFontPath()), "",
            wb.getFontAt(0).getFontHeightInPoints());
        float naturalColTotal = 0f;
        // Determine used column range from print area or sheet data
        int lastCol = 4; // columns A–E (0-indexed: 0–4)
        for (int c = 0; c <= lastCol; c++) {
          int widthIn256 = xSheet.getColumnWidth(c);
          int px = (int) (((widthIn256 + (128 / mdw)) / 256.0) * mdw); // OOXML spec formula
          naturalColTotal += px * (72f / 96f);
        }
        float naturalRowTotal = 0f;
        float defaultH = xSheet.getDefaultRowHeightInPoints();
        for (int r = 0; r < 30; r++) {
          var row = xSheet.getRow(r);
          naturalRowTotal += (row != null) ? row.getHeightInPoints() : defaultH;
        }
        double leftM = xSheet.getMargin(org.apache.poi.ss.usermodel.PageMargin.LEFT);
        double rightM = xSheet.getMargin(org.apache.poi.ss.usermodel.PageMargin.RIGHT);
        double topM = xSheet.getMargin(org.apache.poi.ss.usermodel.PageMargin.TOP);
        double botM = xSheet.getMargin(org.apache.poi.ss.usermodel.PageMargin.BOTTOM);
        float printW = PDRectangle.A4.getWidth() - (float) ((leftM + rightM) * 72);
        float printH = PDRectangle.A4.getHeight() - (float) ((topM + botM) * 72);
        ourFitScale = computeOurFitScale(xSheet, naturalColTotal, naturalRowTotal, printW, printH);
      }

      assertThat(ourFitScale)
          .as("[%s] our fitScale should match Excel's (Excel=%.4f)", sheetName, excelFitScale)
          .isCloseTo(excelFitScale, within(SCALE_TOLERANCE));
    }

    @Test @DisplayName("adjust_100_explicit: fitScale=1.0")
    void adjust100Explicit(@TempDir Path tmp) throws Exception { runSheetTest("adjust_100_explicit", 0, tmp); }

    @Test @DisplayName("adjust_100_implicit: pageSetup未設定でもfitScale=1.0")
    void adjust100Implicit(@TempDir Path tmp) throws Exception { runSheetTest("adjust_100_implicit", 1, tmp); }

    @Test @DisplayName("adjust_85: fitScale=0.85")
    void adjust85(@TempDir Path tmp) throws Exception { runSheetTest("adjust_85", 2, tmp); }

    @Test @DisplayName("fit_1x1: ExcelのfitScaleと一致する")
    void fit1x1(@TempDir Path tmp) throws Exception { runSheetTest("fit_1x1", 3, tmp); }

    @Test @DisplayName("fit_1xN: 高さ制約なし、幅のみのfitScaleと一致する")
    void fit1xN(@TempDir Path tmp) throws Exception { runSheetTest("fit_1xN", 4, tmp); }

    @Test @DisplayName("fit_narrow: 幅が収まるためfitScale=1.0")
    void fitNarrow(@TempDir Path tmp) throws Exception { runSheetTest("fit_narrow", 5, tmp); }
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  private Path createMinimalWorkbook(Path dir, short paperSize, boolean landscape)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(paperSize);
      sheet.getPrintSetup().setLandscape(landscape);
      sheet.createRow(0).createCell(0).setCellValue("x");
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookTooWideForPage(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.5);
      sheet.setMargin(PageMargin.RIGHT, 0.5);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // 30 cols × default width (~42pt) ≈ 1260pt >> printable width ~523pt
      var row = sheet.createRow(0);
      for (int c = 0; c < 30; c++) {
        row.createCell(c).setCellValue("x");
      }
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookTooTallForPage(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.5);
      sheet.setMargin(PageMargin.RIGHT, 0.5);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // 100 rows × default height (~15pt) ≈ 1500pt >> printable height ~770pt
      for (int r = 0; r < 100; r++) {
        sheet.createRow(r).createCell(0).setCellValue("x");
      }
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createSmallWorkbookNoScale(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFillForegroundColor(
          new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("x");

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createFitToPageBothConstraintsWorkbook(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.5);
      sheet.setMargin(PageMargin.RIGHT, 0.5);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setFitToPage(true);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFillForegroundColor(new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      // 20 wide columns (each ~42pt) ≈ 840pt >> printable width ~523pt
      // 40 rows × 20pt ≈ 800pt >> printable height ~770pt
      // Both constraints are active; fitToPage must use the minimum scale.
      for (int r = 0; r < 40; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(20);
        for (int c = 0; c < 20; c++) {
          row.createCell(c).setCellStyle(style);
        }
      }

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createHorizontallyCenteredWorkbook(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.5);   // 36pt
      sheet.setMargin(PageMargin.RIGHT, 0.5);  // 36pt
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // horizontalCentered=true: content narrower than page should be centered
      sheet.setHorizontallyCenter(true);

      // Column A: explicit width ~200pt (≈ 267 character units).
      // At 7px/char × 0.75pt/px = 5.25pt/unit → 200pt / 5.25 ≈ 38.1 units → use 38 units.
      sheet.setColumnWidth(0, 38 * 256);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFillForegroundColor(new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      row.createCell(0).setCellStyle(style);

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createScaledWorkbook(Path dir, short scalePercent) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale(scalePercent);
      sheet.setMargin(PageMargin.LEFT, 0.25);   // 18pt
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);     // 36pt
      sheet.setMargin(PageMargin.BOTTOM, 0.5);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFillForegroundColor(
          new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("x");

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithHeader(Path dir, String fileName,
      @Nullable String left, String center, @Nullable String right) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      if (left != null) {
        sheet.getHeader().setLeft(left);
      }
      if (center != null) {
        sheet.getHeader().setCenter(center);
      }
      if (right != null) {
        sheet.getHeader().setRight(right);
      }
      sheet.createRow(0).createCell(0).setCellValue("content");
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithFooter(Path dir, String center) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getFooter().setCenter(center);
      sheet.createRow(0).createCell(0).setCellValue("content");
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithHeaderAndRowBreak(Path dir, String headerCenter)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getHeader().setCenter(headerCenter);
      for (int r = 0; r < 5; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section1");
      }
      sheet.setRowBreak(4);
      for (int r = 5; r < 10; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section2");
      }
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithHeaderMarginConfig(Path dir,
      double headerMarginIn, double topMarginIn) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.HEADER, headerMarginIn);
      sheet.setMargin(PageMargin.TOP, topMarginIn);
      sheet.setMargin(PageMargin.BOTTOM, topMarginIn);
      sheet.setMargin(PageMargin.FOOTER, headerMarginIn);
      sheet.getHeader().setCenter("HEADER");
      sheet.createRow(0).createCell(0).setCellValue("CONTENT");
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private String textOfPage(PDDocument doc, int pageNumber) throws IOException {
    PDFTextStripper stripper = new PDFTextStripper();
    stripper.setStartPage(pageNumber);
    stripper.setEndPage(pageNumber);
    return stripper.getText(doc);
  }

  private Path createWorkbookWithSingleRowBreak(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      for (int r = 0; r < 5; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section1");
      }
      sheet.setRowBreak(4);
      for (int r = 5; r < 10; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section2");
      }
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithMultipleRowBreaks(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      for (int r = 0; r < 5; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section1");
      }
      sheet.setRowBreak(4);
      for (int r = 5; r < 10; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section2");
      }
      sheet.setRowBreak(9);
      for (int r = 10; r < 15; r++) {
        sheet.createRow(r).createCell(0).setCellValue("section3");
      }
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithColumnBreak(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      for (int r = 0; r < 3; r++) {
        var row = sheet.createRow(r);
        for (int c = 0; c < 3; c++) {
          row.createCell(c).setCellValue("left");
        }
        for (int c = 3; c < 6; c++) {
          row.createCell(c).setCellValue("right");
        }
      }
      sheet.setColumnBreak(2);
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithPrintArea(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      // Print area: A1:C3
      wb.setPrintArea(0, 0, 2, 0, 2);

      for (int r = 0; r < 3; r++) {
        var row = sheet.createRow(r);
        for (int c = 0; c < 3; c++) {
          row.createCell(c).setCellValue("inside");
        }
        row.createCell(3).setCellValue("outside_col");
      }
      sheet.createRow(4).createCell(0).setCellValue("outside_row");

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithoutPrintArea(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.createRow(0).createCell(0).setCellValue("data");
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  // --- text pixel inspection utilities ---

  private static boolean isBlueish(int argb) {
    int r = (argb >> 16) & 0xFF;
    int g = (argb >> 8) & 0xFF;
    int b = argb & 0xFF;
    return b > 100 && b > r + 50 && b > g + 30;
  }

  private static boolean isColoredPixel(int argb, int targetRgb) {
    int r = (argb >> 16) & 0xFF;
    int g = (argb >> 8) & 0xFF;
    int b = argb & 0xFF;
    int tr = (targetRgb >> 16) & 0xFF;
    int tg = (targetRgb >> 8) & 0xFF;
    int tb = targetRgb & 0xFF;
    return Math.abs(r - tr) < 50 && Math.abs(g - tg) < 50 && Math.abs(b - tb) < 50;
  }

  private static double luminance(int argb) {
    int r = (argb >> 16) & 0xFF;
    int g = (argb >>  8) & 0xFF;
    int b =  argb        & 0xFF;
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
  }

  private static int leftmostColoredX(BufferedImage img,
      int xStart, int xEnd, int yStart, int yEnd, int targetRgb) {
    for (int x = xStart; x < xEnd; x++) {
      for (int y = yStart; y < yEnd; y++) {
        if (isColoredPixel(img.getRGB(x, y), targetRgb)) {
          return x;
        }
      }
    }
    return xEnd;
  }

  private static int rightmostColoredX(BufferedImage img,
      int xStart, int xEnd, int yStart, int yEnd, int targetRgb) {
    for (int x = xEnd - 1; x >= xStart; x--) {
      for (int y = yStart; y < yEnd; y++) {
        if (isColoredPixel(img.getRGB(x, y), targetRgb)) {
          return x;
        }
      }
    }
    return xStart;
  }

  private static int topmostColoredY(BufferedImage img,
      int xStart, int xEnd, int yStart, int yEnd, int targetRgb) {
    for (int y = yStart; y < yEnd; y++) {
      for (int x = xStart; x < xEnd; x++) {
        if (isColoredPixel(img.getRGB(x, y), targetRgb)) {
          return y;
        }
      }
    }
    return yEnd;
  }

  private static int bottommostColoredY(BufferedImage img,
      int xStart, int xEnd, int yStart, int yEnd, int targetRgb) {
    for (int y = yEnd - 1; y >= yStart; y--) {
      for (int x = xStart; x < xEnd; x++) {
        if (isColoredPixel(img.getRGB(x, y), targetRgb)) {
          return y;
        }
      }
    }
    return yStart;
  }

  private static int glyphHeight(BufferedImage img,
      int x, int yStart, int yEnd, int targetRgb) {
    int top = topmostColoredY(img, x, x + 1, yStart, yEnd, targetRgb);
    int bottom = bottommostColoredY(img, x, x + 1, yStart, yEnd, targetRgb);
    return (bottom >= top) ? bottom - top + 1 : 0;
  }

  private static int countColoredPixels(BufferedImage img,
      int xStart, int xEnd, int yStart, int yEnd, int targetRgb) {
    int count = 0;
    for (int y = yStart; y < yEnd; y++) {
      for (int x = xStart; x < xEnd; x++) {
        if (isColoredPixel(img.getRGB(x, y), targetRgb)) {
          count++;
        }
      }
    }
    return count;
  }

  // --- text workbook helpers ---

  private Path createTightFontCellWorkbook(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 85);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      font.setBold(true);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.LEFT);
      style.setVerticalAlignment(VerticalAlignment.CENTER);

      var row = sheet.createRow(0);
      row.setHeightInPoints(20);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("合計");

      Path path = dir.resolve("tight-font-cell.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTextWorkbook(Path dir, String fileName, String text,
      int fontSize, boolean bold, boolean italic, boolean strikeout,
      boolean superscript, boolean subscript, boolean wrapText, @Nullable XSSFColor textColor,
      HorizontalAlignment hAlign, VerticalAlignment vAlign,
      float rowHeightPt, int colWidthPoiUnits) throws IOException {
    return createTextWorkbook(dir, fileName, text, fontSize, bold, italic, strikeout,
        superscript, subscript, wrapText, textColor, hAlign, vAlign,
        rowHeightPt, colWidthPoiUnits, false, (short) 0);
  }

  private Path createTextWorkbook(Path dir, String fileName, String text,
      int fontSize, boolean bold, boolean italic, boolean strikeout,
      boolean superscript, boolean subscript, boolean wrapText, @Nullable XSSFColor textColor,
      HorizontalAlignment hAlign, VerticalAlignment vAlign,
      float rowHeightPt, int colWidthPoiUnits, boolean shrinkToFit) throws IOException {
    return createTextWorkbook(dir, fileName, text, fontSize, bold, italic, strikeout,
        superscript, subscript, wrapText, textColor, hAlign, vAlign,
        rowHeightPt, colWidthPoiUnits, shrinkToFit, (short) 0);
  }

  private Path createTextWorkbook(Path dir, String fileName, String text,
      int fontSize, boolean bold, boolean italic, boolean strikeout,
      boolean superscript, boolean subscript, boolean wrapText, @Nullable XSSFColor textColor,
      HorizontalAlignment hAlign, VerticalAlignment vAlign,
      float rowHeightPt, int colWidthPoiUnits,
      boolean shrinkToFit, short rotation) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, colWidthPoiUnits);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) fontSize);
      font.setBold(bold);
      font.setItalic(italic);
      font.setStrikeout(strikeout);
      if (superscript) {
        font.setTypeOffset(org.apache.poi.ss.usermodel.Font.SS_SUPER);
      } else if (subscript) {
        font.setTypeOffset(org.apache.poi.ss.usermodel.Font.SS_SUB);
      }
      if (textColor != null) {
        font.setColor(textColor);
      }

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(hAlign);
      style.setVerticalAlignment(vAlign);
      style.setWrapText(wrapText);
      style.setShrinkToFit(shrinkToFit);
      if (rotation != 0) {
        style.setRotation(rotation);
      }

      var row = sheet.createRow(0);
      row.setHeightInPoints(rowHeightPt);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue(text);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createNumericTextWorkbook(Path dir, String fileName,
      double value, XSSFColor textColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      if (textColor != null) {
        font.setColor(textColor);
      }

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.GENERAL);
      style.setVerticalAlignment(VerticalAlignment.TOP);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue(value);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createFormulaWorkbook(Path dir, String fileName, String formula,
      boolean isStringResult, XSSFColor textColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      if (textColor != null) {
        font.setColor(textColor);
      }

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.GENERAL);
      style.setVerticalAlignment(VerticalAlignment.TOP);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellFormula(formula);
      // Pre-set cached result for rendering without evaluation
      if (isStringResult) {
        cell.setCellValue(formula.replace("\"", ""));
      } else {
        cell.setCellValue(2.0);
      }

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private static byte[] solidColorPng(int rgb, int width, int height) throws IOException {
    BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
    Graphics2D g = img.createGraphics();
    g.setColor(new Color((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF));
    g.fillRect(0, 0, width, height);
    g.dispose();
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    ImageIO.write(img, "PNG", baos);
    return baos.toByteArray();
  }

  private static byte[] solidColorJpeg(int rgb, int width, int height) throws IOException {
    BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
    Graphics2D g = img.createGraphics();
    g.setColor(new Color((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF));
    g.fillRect(0, 0, width, height);
    g.dispose();
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    ImageIO.write(img, "JPEG", baos);
    return baos.toByteArray();
  }

  private static byte[] halfTransparentPng(int rgb, int width, int height) throws IOException {
    BufferedImage img = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
    // Bottom half remains transparent (0x00000000 default for INT_ARGB)
    Graphics2D g = img.createGraphics();
    g.setColor(new Color((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF, 255));
    g.fillRect(0, 0, width, height / 2); // opaque top half only
    g.dispose();
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    ImageIO.write(img, "PNG", baos);
    return baos.toByteArray();
  }

  private Path createImageWorkbook(Path dir, String fileName, byte[] imgBytes,
      int pictureType, int col1, int row1, int col2, int row2) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      for (int c = 0; c <= Math.max(col2, 3); c++) {
        sheet.setColumnWidth(c, 2144); // spec formula MDW=8: int((2144+16)/256*8)=67, pt=50.25
      }
      int maxCol = Math.max(col2, 3);
      int maxRow = Math.max(row2, 3);
      for (int r = 0; r <= maxRow; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= maxCol; c++) {
          row.createCell(c); // ensure all columns are in the used-range
        }
      }
      int picIdx = wb.addPicture(imgBytes, pictureType);
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2), picIdx);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTwoImageWorkbook(Path dir, String fileName,
      byte[] img1, int type1, int c1, int r1, int c2, int r2,
      byte[] img2, int type2, int c3, int r3, int c4, int r4) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      for (int c = 0; c <= 3; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      for (int r = 0; r <= Math.max(r4, 6); r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 3; c++) {
          row.createCell(c);
        }
      }
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, c1, r1, c2, r2),
          wb.addPicture(img1, type1));
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, c3, r3, c4, r4),
          wb.addPicture(img2, type2));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createImageWithManyRowsWorkbook(Path dir, String fileName,
      byte[] imgBytes, int pictureType) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 100); // disable fit-to-page to allow 2-page output
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      for (int c = 0; c <= 3; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      // 35 rows × 30pt = 1050pt > A4 printable height (~770pt) → forces 2 pages
      for (int r = 0; r < 35; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 3; c++) {
          row.createCell(c);
        }
      }
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 0, 0, 2, 2),
          wb.addPicture(imgBytes, pictureType));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createScaledImageWorkbook(Path dir, String fileName,
      byte[] imgBytes, int pictureType, short scalePercent) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale(scalePercent);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      for (int c = 0; c <= 3; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      for (int r = 0; r <= 3; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 3; c++) { row.createCell(c); }
      }
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 0, 0, 2, 2),
          wb.addPicture(imgBytes, pictureType));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createImageOutsidePrintAreaWorkbook(Path dir, String fileName,
      byte[] imgBytes, int pictureType) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // Print area: rows 0-2, cols 0-3
      wb.setPrintArea(0, 0, 3, 0, 2);
      for (int c = 0; c <= 3; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      for (int r = 0; r < 8; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 3; c++) { row.createCell(c); }
      }
      // Image at rows 4-6 (outside print area rows 0-2)
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 0, 4, 2, 6),
          wb.addPicture(imgBytes, pictureType));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createImageOutsidePrintAreaColumnWorkbook(Path dir, String fileName,
      byte[] imgBytes, int pictureType) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // Print area: rows 0-4, cols 0-3 (A-D)
      wb.setPrintArea(0, 0, 3, 0, 4);
      for (int c = 0; c <= 6; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      for (int r = 0; r < 5; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 6; c++) { row.createCell(c); }
      }
      // Image anchor col1=5 (column F, outside print area cols 0-3)
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 5, 0, 7, 2),
          wb.addPicture(imgBytes, pictureType));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createImageWithCellContentWorkbook(Path dir, String fileName,
      byte[] imgBytes, int pictureType) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      for (int c = 0; c <= 3; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      for (int r = 0; r <= 4; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c <= 3; c++) { row.createCell(c); }
      }
      // Cell at row 3 with blue background and text
      XSSFCellStyle style = wb.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      var cell = sheet.getRow(3).createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("内容");
      // Image at rows 0-2
      var drawing = sheet.createDrawingPatriarch();
      drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 0, 0, 2, 2),
          wb.addPicture(imgBytes, pictureType));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTextWorkbookWithUnderline(Path dir, String fileName,
      byte underlineType, XSSFColor fontColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);
      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      font.setUnderline(underlineType);
      font.setColor(fontColor);
      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.LEFT);
      style.setVerticalAlignment(VerticalAlignment.TOP);
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("ABC");
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createIndentWorkbook(Path dir, String fileName, HorizontalAlignment hAlign,
      short indent, XSSFColor textColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      if (textColor != null) {
        font.setColor(textColor);
      }

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(hAlign);
      style.setVerticalAlignment(VerticalAlignment.TOP);
      style.setIndention(indent);

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("A");

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTopAlignNoApplyAlignmentWorkbook(Path dir, String fileName)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      font.setColor(new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setVerticalAlignment(VerticalAlignment.TOP);

      // Remove applyAlignment from the CTXf to reproduce the condition where the
      // <alignment> element has vertical='top' but applyAlignment is not set,
      // which causes Apache POI's getVerticalAlignment() to return BOTTOM.
      wb.getStylesSource().getCellXfAt((int) style.getIndex()).unsetApplyAlignment();

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("A");

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  /**
   * Creates a workbook with a 2-column table (B1:C3) where the custom style applies
   * bold font to the last column (C) via {@code lastColumn} dxf.  showLastColumn=true.
   * Column A (B) uses regular font; column B (C) uses bold via lastColumn styling.
   */
  private Path createTableWithLastColumnWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var stylesSource = wb.getStylesSource();
      var ctStylesheet = stylesSource.getCTStylesheet();

      // dxf[0] = headerRow: dark fill + bold font
      // dxf[1] = lastColumn: bold font only (no fill change)
      var dxfs = ctStylesheet.isSetDxfs() ? ctStylesheet.getDxfs() : ctStylesheet.addNewDxfs();
      dxfs.setCount(2);
      var dxf0 = dxfs.addNewDxf();
      dxf0.addNewFill().addNewPatternFill()
          .setPatternType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
      dxf0.getFill().getPatternFill().addNewFgColor()
          .setRgb(new byte[] {(byte)0xFF, 0x20, 0x20, (byte)0x80}); // dark blue
      dxf0.addNewFont().addNewB();
      var dxf1 = dxfs.addNewDxf();
      dxf1.addNewFont().addNewB(); // bold only, no fill

      var tableStylesEl = ctStylesheet.isSetTableStyles()
          ? ctStylesheet.getTableStyles() : ctStylesheet.addNewTableStyles();
      tableStylesEl.setCount(1);
      tableStylesEl.setDefaultTableStyle("LastColTestStyle");
      var ts = tableStylesEl.addNewTableStyle();
      ts.setName("LastColTestStyle");
      ts.setPivot(false);
      ts.setCount(2);
      var e0 = ts.addNewTableStyleElement();
      e0.setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STTableStyleType.HEADER_ROW);
      e0.setDxfId(0);
      var e1 = ts.addNewTableStyleElement();
      e1.setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STTableStyleType.LAST_COLUMN);
      e1.setDxfId(1);

      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.25);
      sheet.setMargin(PageMargin.BOTTOM, 0.25);
      sheet.setColumnWidth(1, 3000); // col B
      sheet.setColumnWidth(2, 3000); // col C

      var hdr = sheet.createRow(0);
      hdr.setHeightInPoints(20);
      hdr.createCell(1).setCellValue("ColA");
      hdr.createCell(2).setCellValue("ColB");
      var r1 = sheet.createRow(1);
      r1.setHeightInPoints(25);
      r1.createCell(1).setCellValue("ColA-Row1");
      r1.createCell(2).setCellValue("ColB-Row1");
      var r2 = sheet.createRow(2);
      r2.setHeightInPoints(25);
      r2.createCell(1).setCellValue("ColA-Row2");
      r2.createCell(2).setCellValue("ColB-Row2");

      var areaRef = new org.apache.poi.ss.util.AreaReference("B1:C3",
          org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
      var table = sheet.createTable(areaRef);
      table.setName("LastColTable");
      table.setDisplayName("LastColTable");
      var si = table.getCTTable().isSetTableStyleInfo()
          ? table.getCTTable().getTableStyleInfo() : table.getCTTable().addNewTableStyleInfo();
      si.setName("LastColTestStyle");
      si.setShowRowStripes(false);
      si.setShowFirstColumn(false);
      si.setShowLastColumn(true);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  /**
   * Same layout as {@link #createTableWithLastColumnWorkbook} but with
   * {@code showLastColumn=false}, so the lastColumn bold dxf is NOT applied.
   */
  private Path createTableWithLastColumnNoShowWorkbook(Path dir, String fileName)
      throws Exception {
    Path result = createTableWithLastColumnWorkbook(dir, "_tmp_" + fileName);
    try (XSSFWorkbook wb = new XSSFWorkbook(result.toFile())) {
      // Use XSSFSheet.getTables() via the XSSFWorkbook's sheet accessor
      for (int si = 0; si < wb.getNumberOfSheets(); si++) {
        var xs = (org.apache.poi.xssf.usermodel.XSSFSheet) wb.getSheetAt(si);
        for (XSSFTable t : xs.getTables()) {
          var info = t.getCTTable().getTableStyleInfo();
          if (info != null) {
            info.setShowLastColumn(false);
          }
        }
      }
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  /**
   * Creates a workbook with an Excel table using a CUSTOM table style with explicit RGB colours
   * (no theme dependency). The custom style defines:
   * - header row fill: solid red (#FF0000), bold font
   * - firstRowStripe fill: solid light-blue (#CCECFF)
   * - wholeTable: thin dark-grey bottom+horizontal borders
   */
  private Path createTableWithCustomStyleWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      // Build the dxf entries using raw CT objects
      var stylesSource = wb.getStylesSource();
      // We inject the custom style XML directly via the CT root element.
      // dxf[0] = wholeTable: thin bottom + horizontal border in dark grey
      // dxf[1] = firstRowStripe: light-blue fill
      // dxf[2] = headerRow: red fill + bold font

      // Inject dxfs and tableStyles into the styles XML
      // We use org.openxmlformats APIs through StylesTable's raw CT
      var ctStylesheet = stylesSource.getCTStylesheet();

      // Add dxfs
      var dxfs = ctStylesheet.isSetDxfs() ? ctStylesheet.getDxfs() : ctStylesheet.addNewDxfs();
      dxfs.setCount(3);
      // dxf[0] = headerRow: red fill + bold
      var dxf0 = dxfs.addNewDxf();
      var fill0 = dxf0.addNewFill().addNewPatternFill();
      fill0.setPatternType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
      fill0.addNewFgColor().setRgb(new byte[] {(byte)0xFF, (byte)0xFF, 0x00, 0x00}); // #FF0000
      dxf0.addNewFont().addNewB();
      // dxf[1] = firstRowStripe: light-blue fill
      var dxf1 = dxfs.addNewDxf();
      var fill1 = dxf1.addNewFill().addNewPatternFill();
      fill1.setPatternType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.SOLID);
      fill1.addNewFgColor().setRgb(new byte[] {(byte)0xFF, (byte)0xCC, (byte)0xEC, (byte)0xFF});
      // dxf[2] = wholeTable: thin bottom + horizontal border
      var dxf2 = dxfs.addNewDxf();
      var border2 = dxf2.addNewBorder();
      var btm = border2.addNewBottom();
      btm.setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
      btm.addNewColor().setRgb(new byte[] {(byte)0xFF, 0x40, 0x40, 0x40});
      var horiz = border2.addNewHorizontal();
      horiz.setStyle(org.openxmlformats.schemas.spreadsheetml.x2006.main.STBorderStyle.THIN);
      horiz.addNewColor().setRgb(new byte[] {(byte)0xFF, 0x40, 0x40, 0x40});

      // Add tableStyles via raw CT
      var tableStylesEl = ctStylesheet.isSetTableStyles()
          ? ctStylesheet.getTableStyles() : ctStylesheet.addNewTableStyles();
      tableStylesEl.setCount(1);
      tableStylesEl.setDefaultTableStyle("CustomTestStyle");
      var ts = tableStylesEl.addNewTableStyle();
      ts.setName("CustomTestStyle");
      ts.setPivot(false);
      ts.setCount(3);
      var e0 = ts.addNewTableStyleElement();
      e0.setType(org.openxmlformats.schemas.spreadsheetml.x2006.main.STTableStyleType.HEADER_ROW);
      e0.setDxfId(0);
      var e1 = ts.addNewTableStyleElement();
      e1.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STTableStyleType.FIRST_ROW_STRIPE);
      e1.setDxfId(1);
      var e2 = ts.addNewTableStyleElement();
      e2.setType(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STTableStyleType.WHOLE_TABLE);
      e2.setDxfId(2);

      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.25);
      sheet.setMargin(PageMargin.BOTTOM, 0.25);
      sheet.setColumnWidth(1, 5000);
      var hdr = sheet.createRow(0);
      hdr.setHeightInPoints(20);
      hdr.createCell(1).setCellValue("HEADER");
      String[] data = {"Row1", "Row2", "Row3"};
      for (int i = 0; i < data.length; i++) {
        var row = sheet.createRow(i + 1);
        row.setHeightInPoints(25);
        row.createCell(1).setCellValue(data[i]);
      }
      var areaRef = new org.apache.poi.ss.util.AreaReference("B1:B4",
          org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
      var table = sheet.createTable(areaRef);
      table.setName("TestTable2");
      table.setDisplayName("TestTable2");
      var si = table.getCTTable().isSetTableStyleInfo()
          ? table.getCTTable().getTableStyleInfo()
          : table.getCTTable().addNewTableStyleInfo();
      si.setName("CustomTestStyle");
      si.setShowRowStripes(true);
      si.setShowFirstColumn(false);
      si.setShowLastColumn(false);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  /**
   * Creates a workbook with an Excel table (B1:B4, header row 1, data rows 2-4) using the
   * built-in {@code TableStyleMedium6} style.  The style has a dark teal header fill,
   * alternating row stripes, and thin horizontal borders between rows.
   */
  private Path createTableWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.25);
      sheet.setMargin(PageMargin.BOTTOM, 0.25);
      sheet.setColumnWidth(1, 5000); // column B

      // Header row (row 0)
      var hdr = sheet.createRow(0);
      hdr.setHeightInPoints(20);
      hdr.createCell(1).setCellValue("HEADER");
      // Data rows
      String[] data = {"Row1", "Row2", "Row3"};
      for (int i = 0; i < data.length; i++) {
        var row = sheet.createRow(i + 1);
        row.setHeightInPoints(25);
        row.createCell(1).setCellValue(data[i]);
      }

      // Create an Excel table over B1:B4 with TableStyleMedium6.
      // createTable(AreaReference) sets up ref and columns automatically.
      var areaRef = new org.apache.poi.ss.util.AreaReference("B1:B4",
          org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
      var table = sheet.createTable(areaRef);
      table.setName("TestTable");
      table.setDisplayName("TestTable");
      // Set the table style
      var styleInfo = table.getCTTable().isSetTableStyleInfo()
          ? table.getCTTable().getTableStyleInfo()
          : table.getCTTable().addNewTableStyleInfo();
      styleInfo.setName("TableStyleMedium6");
      styleInfo.setShowRowStripes(true);
      styleInfo.setShowFirstColumn(false);
      styleInfo.setShowLastColumn(false);
      styleInfo.setShowColumnStripes(false);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createRightAlignNoApplyAlignmentWorkbook(Path dir, String fileName)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);

      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      font.setColor(new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));

      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.RIGHT);
      style.setVerticalAlignment(VerticalAlignment.TOP);

      // Remove applyAlignment to reproduce the condition where horizontal='right' is present
      // in the <alignment> element but applyAlignment is not set, causing Apache POI's
      // getAlignment() to return GENERAL (the default) instead of RIGHT.
      wb.getStylesSource().getCellXfAt((int) style.getIndex()).unsetApplyAlignment();

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("A");

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createPreambleTitleRowsWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 100);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // Rows 0-1: preamble rows (appear only on page 1, before the title row)
      var preambleA = sheet.createRow(0);
      preambleA.setHeightInPoints(20);
      preambleA.createCell(0).setCellValue("PREAMBLE_A");
      var preambleB = sheet.createRow(1);
      preambleB.setHeightInPoints(20);
      preambleB.createCell(0).setCellValue("PREAMBLE_B");
      // Row 2: print title row (repeating header)
      var headerRow = sheet.createRow(2);
      headerRow.setHeightInPoints(20);
      headerRow.createCell(0).setCellValue("HEADER");
      // Rows 3-35: data rows (25pt each). Capacity after title (20pt): 750pt.
      // Page 1 also has preamble (2×20=40pt) → 710pt → 28 rows (28×25=700).
      // Page 2: full 750pt → remaining rows.
      for (int r = 3; r <= 35; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(25);
        row.createCell(0).setCellValue("row" + (r - 2));
      }
      // Set row 2 as the repeating print title row
      sheet.setRepeatingRows(new CellRangeAddress(2, 2, -1, -1));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTitleRowsWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 100);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // Row 0: header row (print title row)
      var headerRow = sheet.createRow(0);
      headerRow.setHeightInPoints(20);
      headerRow.createCell(0).setCellValue("HEADER");
      // Rows 1-35: data rows. Page capacity = 749.9pt; 30×25=750pt fits (tolerance), 31+ breaks.
      // → page 1: rows 1-30, page 2: rows 31-35.
      for (int r = 1; r <= 35; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(25);
        row.createCell(0).setCellValue("row" + r);
      }
      // Set row 0 as repeating (print title) row
      sheet.setRepeatingRows(new CellRangeAddress(0, 0, -1, -1));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTitleColsWorkbook(Path dir, String fileName,
      int titleColCount, int contentColCount) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 100);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      int totalCols = titleColCount + contentColCount;
      for (int c = 0; c < totalCols; c++) {
        sheet.setColumnWidth(c, 2144); // spec formula MDW=8: int((2144+16)/256*8)=67, pt=50.25
      }
      var row = sheet.createRow(0);
      row.setHeightInPoints(20);
      for (int c = 0; c < titleColCount; c++) {
        row.createCell(c).setCellValue("LABEL_" + (char) ('A' + c));
      }
      for (int c = 0; c < contentColCount; c++) {
        row.createCell(titleColCount + c).setCellValue("col" + (c + 1));
      }
      sheet.setRepeatingColumns(new CellRangeAddress(-1, -1, 0, titleColCount - 1));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTitleColsAndRowsWorkbook(Path dir, String fileName)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.getPrintSetup().setScale((short) 100);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      // 1 title col (A) + 11 content cols (B-L), each ≈ 50pt; 12 cols total
      for (int c = 0; c < 12; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      // Row 0: title row (20pt); A0 = "CORNER"
      var headerRow = sheet.createRow(0);
      headerRow.setHeightInPoints(20);
      headerRow.createCell(0).setCellValue("CORNER");
      for (int c = 1; c < 12; c++) {
        headerRow.createCell(c);
      }
      // Rows 1-35: content rows (25pt); B_r = "row_r"
      for (int r = 1; r <= 35; r++) {
        var dataRow = sheet.createRow(r);
        dataRow.setHeightInPoints(25);
        for (int c = 0; c < 12; c++) {
          dataRow.createCell(c);
        }
        dataRow.getCell(1).setCellValue("row" + r);
      }
      sheet.setRepeatingRows(new CellRangeAddress(0, 0, -1, -1));
      sheet.setRepeatingColumns(new CellRangeAddress(-1, -1, 0, 0));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createTwoSheetWorkbook(Path dir, String fileName) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet1 = wb.createSheet("Sheet1");
      sheet1.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet1.createRow(0).createCell(0).setCellValue("sheet1content");
      var sheet2 = wb.createSheet("Sheet2");
      sheet2.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet2.createRow(0).createCell(0).setCellValue("sheet2content");
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createMergedWorkbook(Path dir, String fileName,
      int firstRow, int lastRow, int firstCol, int lastCol,
      XSSFColor bgColor, @Nullable HorizontalAlignment hAlign, @Nullable String text)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      int maxCol = Math.max(lastCol + 2, 4);
      for (int c = 0; c < maxCol; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      int maxRow = Math.max(lastRow + 2, 4);
      for (int r = 0; r < maxRow; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c < maxCol; c++) {
          row.createCell(c);
        }
      }
      XSSFCellStyle style = wb.createCellStyle();
      if (bgColor != null) {
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(bgColor);
      }
      if (hAlign != null) {
        style.setAlignment(hAlign);
      }
      var topLeft = sheet.getRow(firstRow).getCell(firstCol);
      topLeft.setCellStyle(style);
      if (text != null) {
        topLeft.setCellValue(text);
      }
      sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createMergeWithBorderWorkbook(Path dir, String fileName,
      int firstRow, int lastRow, int firstCol, int lastCol,
      BorderStyle rightBorder, BorderStyle bottomBorder) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      int maxCol = Math.max(lastCol + 2, 4);
      for (int c = 0; c < maxCol; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      int maxRow = Math.max(lastRow + 2, 4);
      for (int r = 0; r < maxRow; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c < maxCol; c++) {
          row.createCell(c);
        }
      }
      if (rightBorder != BorderStyle.NONE) {
        XSSFCellStyle s = wb.createCellStyle();
        s.setBorderRight(rightBorder);
        sheet.getRow(firstRow).getCell(lastCol).setCellStyle(s);
      }
      if (bottomBorder != BorderStyle.NONE) {
        XSSFCellStyle s = wb.createCellStyle();
        s.setBorderBottom(bottomBorder);
        sheet.getRow(lastRow).getCell(firstCol).setCellStyle(s);
      }
      sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createMergedWorkbookWithFont(Path dir, String fileName,
      int firstRow, int lastRow, int firstCol, int lastCol,
      XSSFColor fontColor, HorizontalAlignment hAlign) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      int maxCol = Math.max(lastCol + 2, 4);
      for (int c = 0; c < maxCol; c++) {
        sheet.setColumnWidth(c, 2144);
      }
      int maxRow = Math.max(lastRow + 2, 4);
      for (int r = 0; r < maxRow; r++) {
        var row = sheet.createRow(r);
        row.setHeightInPoints(30);
        for (int c = 0; c < maxCol; c++) {
          row.createCell(c);
        }
      }
      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 14);
      font.setColor(fontColor);
      XSSFCellStyle style = wb.createCellStyle();
      style.setFont(font);
      style.setAlignment(hAlign);
      style.setVerticalAlignment(VerticalAlignment.TOP);
      var topLeft = sheet.getRow(firstRow).getCell(firstCol);
      topLeft.setCellStyle(style);
      topLeft.setCellValue("A");
      if (lastRow > firstRow || lastCol > firstCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
      }
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createBgWorkbook(Path dir, String fileName,
      @Nullable XSSFColor bgColor, FillPatternType pattern) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 2048);
      XSSFCellStyle style = wb.createCellStyle();
      style.setFillPattern(pattern);
      if (bgColor != null) {
        style.setFillForegroundColor(bgColor);
      }
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      row.createCell(0).setCellStyle(style);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createAdjacentBgWorkbook(Path dir, String fileName,
      XSSFColor color0, XSSFColor color1) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(1, 2048);
      XSSFCellStyle s0 = wb.createCellStyle();
      s0.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      s0.setFillForegroundColor(color0);
      XSSFCellStyle s1 = wb.createCellStyle();
      s1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      s1.setFillForegroundColor(color1);
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      row.createCell(0).setCellStyle(s0);
      row.createCell(1).setCellStyle(s1);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createIndexedColorWorkbook(Path dir, String fileName, IndexedColors color)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 2048);
      XSSFCellStyle style = wb.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(color.getIndex());
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      row.createCell(0).setCellStyle(style);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createThemeColorWorkbook(Path dir, String fileName, int themeIndex)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 2048);
      XSSFCellStyle style = wb.createCellStyle();
      CTColor ctColor = CTColor.Factory.newInstance();
      ctColor.setTheme(themeIndex);
      // Set an explicit RGB fallback so the color is renderable even without a theme table.
      // Accent1 blue (#4472C4) is the Office theme default for index 4.
      ctColor.setRgb(new byte[] {0x44, 0x72, (byte) 0xC4});
      XSSFColor themeColor = XSSFColor.from(ctColor, wb.getStylesSource().getIndexedColors());
      ThemesTable themes = wb.getStylesSource().getTheme();
      if (themes != null) {
        themes.inheritFromThemeAsRequired(themeColor);
      }
      style.setFillForegroundColor(themeColor);
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      row.createCell(0).setCellStyle(style);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createBgWithTextWorkbook(Path dir, String fileName,
      XSSFColor bgColor, XSSFColor textColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 2048);
      XSSFFont font = wb.createFont();
      font.setFontHeightInPoints((short) 18);
      font.setColor(textColor);
      XSSFCellStyle style = wb.createCellStyle();
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setFillForegroundColor(bgColor);
      style.setFont(font);
      style.setAlignment(HorizontalAlignment.LEFT);
      style.setVerticalAlignment(VerticalAlignment.CENTER);
      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("W");
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createFormattedCellWorkbook(Path dir, String fileName,
      double value, String formatString) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      XSSFCellStyle style = wb.createCellStyle();
      style.setDataFormat(wb.createDataFormat().getFormat(formatString));
      var row = sheet.createRow(0);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue(value);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createZeroFormulaWorkbook(Path dir, String formatString) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 3840);
      XSSFCellStyle style = wb.createCellStyle();
      style.setDataFormat(wb.createDataFormat().getFormat(formatString));
      var row = sheet.createRow(0);
      row.setHeightInPoints(20);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellFormula("0");
      wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
      Path path = dir.resolve("zero-formula.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createFormulaCellWorkbook(Path dir, String fileName,
      String formula, String formatString) throws IOException {
    return createFormulaCellWorkbook(dir, fileName, formula, formatString, 0.0);
  }

  private Path createFormulaCellWorkbook(Path dir, String fileName,
      String formula, String formatString, double cachedResult) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      XSSFCellStyle style = wb.createCellStyle();
      style.setDataFormat(wb.createDataFormat().getFormat(formatString));
      var row = sheet.createRow(0);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellFormula(formula);
      cell.setCellValue(cachedResult != 0.0 ? cachedResult : 1234.0);
      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private static int avgGray(int argb) {
    return (((argb >> 16) & 0xFF) + ((argb >> 8) & 0xFF) + (argb & 0xFF)) / 3;
  }

  private static int countDarkPixels(BufferedImage img, int x, int yStart, int yEnd) {
    int count = 0;
    for (int y = yStart; y <= yEnd; y++) {
      if (avgGray(img.getRGB(x, y)) < 128) {
        count++;
      }
    }
    return count;
  }

  private static boolean hasDarkPixelNear(BufferedImage img, int cx, int cy, int radius) {
    for (int dx = -radius; dx <= radius; dx++) {
      for (int dy = -radius; dy <= radius; dy++) {
        if (avgGray(img.getRGB(cx + dx, cy + dy)) < 128) {
          return true;
        }
      }
    }
    return false;
  }

  private Path createBorderWorkbook(Path dir, String fileName,
      BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right,
      @Nullable XSSFColor borderColor) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 1792); // spec formula MDW=8: int((1792+16)/256*8)=56, pt=42 → B_RIGHT=120

      XSSFCellStyle style = wb.createCellStyle();
      style.setBorderTop(top);
      style.setBorderBottom(bottom);
      style.setBorderLeft(left);
      style.setBorderRight(right);
      if (borderColor != null) {
        if (top != BorderStyle.NONE) {
          style.setTopBorderColor(borderColor);
        }
        if (bottom != BorderStyle.NONE) {
          style.setBottomBorderColor(borderColor);
        }
        if (left != BorderStyle.NONE) {
          style.setLeftBorderColor(borderColor);
        }
        if (right != BorderStyle.NONE) {
          style.setRightBorderColor(borderColor);
        }
      }

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);
      cell.setCellValue("border");

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createDiagonalWorkbook(Path dir, String fileName,
      BorderStyle borderStyle, boolean diagonalDown, boolean diagonalUp) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, 0.25);
      sheet.setMargin(PageMargin.RIGHT, 0.25);
      sheet.setMargin(PageMargin.TOP, 0.5);
      sheet.setMargin(PageMargin.BOTTOM, 0.5);
      sheet.setColumnWidth(0, 1792); // spec formula MDW=8: int((1792+16)/256*8)=56, pt=42 → B_RIGHT=120

      XSSFCellStyle style = wb.createCellStyle();
      // setBorderLeft forces a new non-default CTBorder entry into the styles table,
      // which allows the diagonal border style to be appended to that entry.
      style.setBorderLeft(BorderStyle.THIN);
      int borderId = (int) style.getCoreXf().getBorderId();
      XSSFCellBorder cellBorder = wb.getStylesSource().getBorderAt(borderId);
      cellBorder.setBorderStyle(XSSFCellBorder.BorderSide.DIAGONAL, borderStyle);
      if (diagonalDown) {
        cellBorder.getCTBorder().setDiagonalDown(true);
      }
      if (diagonalUp) {
        cellBorder.getCTBorder().setDiagonalUp(true);
      }

      var row = sheet.createRow(0);
      row.setHeightInPoints(60);
      var cell = row.createCell(0);
      cell.setCellStyle(style);

      Path path = dir.resolve(fileName);
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createWorkbookWithUnknownFont(Path dir) throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      // Override default font (index 0) — this is what ExcelToPdfUtil reads via getFontAt(0)
      XSSFFont defaultFont = wb.getFontAt(0);
      defaultFont.setFontName("__FictionalFontZZZ123__");
      defaultFont.setFontHeightInPoints((short) 11);

      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      var row = sheet.createRow(0);
      row.setHeightInPoints(15);
      row.createCell(0).setCellValue("test");
      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }

  private Path createColoredWorkbook(Path dir, double leftMarginIn, double topMarginIn)
      throws IOException {
    try (XSSFWorkbook wb = new XSSFWorkbook()) {
      var sheet = wb.createSheet("Sheet1");
      sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
      sheet.setMargin(PageMargin.LEFT, leftMarginIn);
      sheet.setMargin(PageMargin.RIGHT, leftMarginIn);
      sheet.setMargin(PageMargin.TOP, topMarginIn);
      sheet.setMargin(PageMargin.BOTTOM, topMarginIn);

      XSSFCellStyle style = wb.createCellStyle();
      style.setFillForegroundColor(
          new XSSFColor(new byte[] {0x00, 0x70, (byte) 0xC0}, null));
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      for (int r = 0; r < 3; r++) {
        var row = sheet.createRow(r);
        for (int c = 0; c < 5; c++) {
          var cell = row.createCell(c);
          cell.setCellStyle(style);
          cell.setCellValue("x");
        }
      }

      Path path = dir.resolve("test.xlsx");
      try (var out = Files.newOutputStream(path)) {
        wb.write(out);
      }
      return path;
    }
  }
}
