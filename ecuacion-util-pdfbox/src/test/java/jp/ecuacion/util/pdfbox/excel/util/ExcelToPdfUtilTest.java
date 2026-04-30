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
package jp.ecuacion.util.pdfbox.excel.util;

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
import javax.imageio.ImageIO;
import jp.ecuacion.util.pdfbox.excel.exception.PdfGenerateException;
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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Page 1");
      }
    }

    @Test
    @DisplayName("&N is replaced with the total number of pages")
    void totalPages(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "of &N", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("of 1");
      }
    }

    @Test
    @DisplayName("&A is replaced with the sheet name")
    void sheetName(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "Sheet: &A", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Sheet: Sheet1");
      }
    }

    @Test
    @DisplayName("&F is replaced with the file name without extension")
    void fileName(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "myreport.xlsx", null, "File: &F", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("A & B");
      }
    }

    @Test
    @DisplayName("&P+n produces page number plus offset")
    void pageNumberWithOffset(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeader(tempDir, "test.xlsx", null, "Page &P+3", null);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(textOfPage(doc, 1)).contains("Page 4");
      }
    }

    @Test
    @DisplayName("header appears on every page with the correct page number")
    void headerRepeatsOnEveryPage(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookWithHeaderAndRowBreak(tempDir, "Page &P of &N");
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
    @DisplayName("content wider than printable area is auto-scaled to fit one page")
    void tooWide(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookTooWideForPage(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
      }
    }

    @Test
    @DisplayName("content taller than printable area is auto-scaled to fit one page")
    void tooTall(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createWorkbookTooTallForPage(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
      }
    }

    @Test
    @DisplayName("content fitting naturally is rendered at natural size without scaling")
    void fitsNaturally(@TempDir Path tempDir) throws IOException, PdfGenerateException {
      Path excel = createSmallWorkbookNoScale(tempDir);
      Path pdf = tempDir.resolve("out.pdf");

      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getNumberOfPages()).isEqualTo(1);
        BufferedImage image = new PDFRenderer(doc).renderImageWithDPI(0, 72);
        // row height = 60pt, top margin = 36pt → row spans y=[36, 96]
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 30) & 0xFFFFFF).isEqualTo(FILL_RGB);
        assertThat(image.getRGB(SAFE_X, TOP_MARGIN_PX + 70) & 0xFFFFFF).isEqualTo(0xFFFFFF);
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(thinXl, List.of("Sheet1"), thinPdf, null);
      ExcelToPdfUtil.generate(thickXl, List.of("Sheet1"), thickPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
    private static final int T_RIGHT   = 196;  // (18+80)pt × 2
    private static final int T_BOTTOM  = 192;  // (36+60)pt × 2
    private static final int T_SAFE_X  = 116;  // (T_LEFT+T_RIGHT)/2
    private static final int T_SAFE_Y  = 132;  // (T_TOP+T_BOTTOM)/2

    // Expected pixel positions given CELL_PADDING=2pt=4px at 144 DPI.
    private static final int T_PAD_LEFT  = T_LEFT  + 4; // 40
    private static final int T_PAD_RIGHT = T_RIGHT - 4; // 192
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(small, List.of("Sheet1"), smallPdf, null);
      ExcelToPdfUtil.generate(large, List.of("Sheet1"), largePdf, null);

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
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, null);
      ExcelToPdfUtil.generate(bold, List.of("Sheet1"), boldPdf, null);

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
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, null);
      ExcelToPdfUtil.generate(italic, List.of("Sheet1"), italicPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(withStrike, List.of("Sheet1"), strikePdf, null);
      ExcelToPdfUtil.generate(noStrike, List.of("Sheet1"), noStrikePdf, null);

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
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, null);
      ExcelToPdfUtil.generate(superXl, List.of("Sheet1"), superPdf, null);

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
      ExcelToPdfUtil.generate(normal, List.of("Sheet1"), normalPdf, null);
      ExcelToPdfUtil.generate(subXl, List.of("Sheet1"), subPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(leftXl, List.of("Sheet1"), leftPdf, null);
      ExcelToPdfUtil.generate(centerXl, List.of("Sheet1"), centerPdf, null);
      ExcelToPdfUtil.generate(rightXl, List.of("Sheet1"), rightPdf, null);

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
      ExcelToPdfUtil.generate(numXl, List.of("Sheet1"), numPdf, null);
      ExcelToPdfUtil.generate(strXl, List.of("Sheet1"), strPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(topXl, List.of("Sheet1"), topPdf, null);
      ExcelToPdfUtil.generate(centerXl, List.of("Sheet1"), centerPdf, null);
      ExcelToPdfUtil.generate(bottomXl, List.of("Sheet1"), bottomPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(noShrink, List.of("Sheet1"), noShrinkPdf, null);
      ExcelToPdfUtil.generate(withShrink, List.of("Sheet1"), withShrinkPdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(withUl, List.of("Sheet1"), ulPdf, null);
      ExcelToPdfUtil.generate(noUl,   List.of("Sheet1"), noUlPdf, null);

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
      ExcelToPdfUtil.generate(singleUl, List.of("Sheet1"), singlePdf, null);
      ExcelToPdfUtil.generate(doubleUl, List.of("Sheet1"), doublePdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);
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
      assertThatThrownBy(() -> ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null))
          .isInstanceOf(RuntimeException.class)
          .hasMessageContaining("Reiwa");
    }

    private void assertFmt(Path dir, String id, double value, String format, String expected)
        throws IOException, PdfGenerateException {
      Path excel = createFormattedCellWorkbook(dir, id + ".xlsx", value, format);
      Path pdf = dir.resolve(id + ".pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
    @DisplayName("image and cell text/background coexist on the same page")
    void imageAndCellContentCoexist(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      byte[] img = solidColorPng(0xFF0000, 80, 60);
      // Cell at row 3 has blue background and text; image at rows 0-2
      Path excel = createImageWithCellContentWorkbook(tempDir, "test.xlsx", img,
          Workbook.PICTURE_TYPE_PNG);
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);
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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(mergedXl, List.of("Sheet1"), mergedPdf, null);
      ExcelToPdfUtil.generate(singleXl, List.of("Sheet1"), singlePdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
  }

  // ---------------------------------------------------------------------------
  // multiple sheets
  // ---------------------------------------------------------------------------

  @Nested
  @DisplayName("multiple sheets")
  class MultipleSheets {

    @Test
    @DisplayName("two sheets produce two PDF pages in the specified order")
    void twoSheetsProduceTwoPages(@TempDir Path tempDir)
        throws IOException, PdfGenerateException {
      Path excel = createTwoSheetWorkbook(tempDir, "test.xlsx");
      Path pdf = tempDir.resolve("out.pdf");
      ExcelToPdfUtil.generate(excel, List.of("Sheet1", "Sheet2"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet2", "Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
      ExcelToPdfUtil.generate(excel, List.of("Sheet1"), pdf, null);

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
        sheet.setColumnWidth(c, 2438); // ≈50pt
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438); // ≈ 50pt per column
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
        sheet.setColumnWidth(c, 2438);
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
      sheet.setColumnWidth(0, 2048); // 8 chars × 256 ≈ 42pt wide

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
      sheet.setColumnWidth(0, 2048);

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
