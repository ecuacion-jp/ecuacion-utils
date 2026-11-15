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
package jp.ecuacion.util.pdf.word.document.util;

import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.field;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.floatingPicture;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.inlinePicture;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.p;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.pageBreak;
import static jp.ecuacion.util.pdf.word.document.util.DocxBuilder.r;
import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import javax.imageio.ImageIO;
import jp.ecuacion.util.pdf.common.exception.FileFormatNotSupportedException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordIncorrectException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordRequiredException;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.word.document.util.PdfLines.Line;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotation;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationLink;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

/**
 * Tests pictures, password-protected files, tracked changes and other containers, and the
 * links / outline made from tables of contents and headings.
 */
@DisplayName("WordToPdfUtil: pictures, passwords, containers and navigation")
class WordToPdfContentTest extends AbstractWordToPdfTest {

  /** Returns a PNG picture filled with the color. */
  private static byte[] png(Color color) throws Exception {
    BufferedImage image = new BufferedImage(20, 10, BufferedImage.TYPE_INT_RGB);
    Graphics2D g = image.createGraphics();
    g.setColor(color);
    g.fillRect(0, 0, 20, 10);
    g.dispose();
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    ImageIO.write(image, "png", out);
    return out.toByteArray();
  }

  private static int imageCount(Path pdf) throws Exception {
    try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
      int count = 0;
      for (COSName name : doc.getPage(0).getResources().getXObjectNames()) {
        if (doc.getPage(0).getResources().isImageXObject(name)) {
          count++;
        }
      }
      return count;
    }
  }

  @Nested
  @DisplayName("pictures")
  class Pictures {

    @Test
    @DisplayName("an inline picture takes its width in the line, and sits on the baseline")
    void inline() throws Exception {
      Path pdf = generate(new DocxBuilder().png("rIdP", png(Color.RED))
          .add(p(r("前") + inlinePicture("rIdP", 60, 30) + r("後"))));
      Line line = PdfLines.of(pdf).get(0).get(0);

      assertThat(line.leftOf("後")).isCloseTo(LEFT + 10 + 60, TOL);
      assertThat(pixel(pdf, LEFT + 10 + 30, line.baseline() - 15)).isEqualTo(0xFF0000);
      assertThat(pixel(pdf, LEFT + 10 + 30, line.baseline() - 35)).isEqualTo(0xFFFFFF);
      assertThat(pixel(pdf, LEFT + 10 + 30, line.baseline() + 3)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("a line with a picture is as tall as the picture")
    void lineHeight() throws Exception {
      final float base = basePitch();
      List<List<Line>> pages = convert(new DocxBuilder().png("rIdP", png(Color.RED))
          .add(p(r("一"))).add(p(inlinePicture("rIdP", 60, 100) + r("二"))).add(p(r("三"))));

      assertThat(pitch(pages, "一", "二")).isGreaterThan(100f);
      assertThat(pitch(pages, "二", "三")).isCloseTo(base, TOL);
    }

    @Test
    @DisplayName("a picture too wide for the rest of the line moves to the next line")
    void wraps() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().png("rIdP", png(Color.RED))
          .add(p(r(chars(40)) + inlinePicture("rIdP", 100, 20) + r("後"))));

      assertThat(pages.get(0)).hasSize(2);
      assertThat(PdfLines.find(pages, "後").leftOf("後")).isCloseTo(LEFT + 100, TOL);
    }

    @Test
    @DisplayName("a picture in a table cell")
    void inCell() throws Exception {
      String table = "<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>"
          + p(inlinePicture("rIdP", 60, 30) + r("後")) + "</w:tc></w:tr></w:tbl>";
      Path pdf = generate(new DocxBuilder().settings(WORD_2013).png("rIdP", png(Color.BLUE))
          .add(table));
      Line line = PdfLines.of(pdf).get(0).get(0);

      assertThat(pixel(pdf, LEFT + 5.4f + 30, line.baseline() - 15)).isEqualTo(0x0000FF);
    }

    @Test
    @DisplayName("a floating picture is at its offset from the page, without moving the text")
    void floatingFromPage() throws Exception {
      Path pdf = generate(new DocxBuilder().png("rIdP", png(Color.GREEN))
          .add(p(floatingPicture("rIdP", 80, 40, "page", 300, "page", 400) + r("本文"))));
      Line line = PdfLines.of(pdf).get(0).get(0);

      assertThat(line.left()).isCloseTo(LEFT, TOL);
      assertThat(pixel(pdf, 300 + 40, 400 + 20)).isEqualTo(0x00FF00);
      assertThat(pixel(pdf, 300 - 5, 400 + 20)).isEqualTo(0xFFFFFF);
      assertThat(pixel(pdf, 300 + 40, 400 + 45)).isEqualTo(0xFFFFFF);
    }

    @Test
    @DisplayName("a floating picture positioned from the margin and the paragraph")
    void floatingFromMarginAndParagraph() throws Exception {
      // (basePitch() converts another document, so it is called before this one.)
      final float paragraphTop = 72 + basePitch();
      Path pdf = generate(new DocxBuilder().png("rIdP", png(Color.GREEN)).add(p(r("一")))
          .add(p(floatingPicture("rIdP", 80, 40, "margin", 100, "paragraph", 50) + r("二"))));

      assertThat(pixel(pdf, LEFT + 100 + 40, paragraphTop + 50 + 20)).isEqualTo(0x00FF00);
    }

    @Test
    @DisplayName("the same picture used twice is embedded once")
    void embeddedOnce() throws Exception {
      Path pdf = generate(new DocxBuilder().png("rIdP", png(Color.RED))
          .add(p(inlinePicture("rIdP", 60, 30) + inlinePicture("rIdP", 30, 15))));

      assertThat(imageCount(pdf)).isEqualTo(1);
    }

    @Test
    @DisplayName("a picture in a format which cannot be read is skipped, the text is output")
    void unreadablePicture() throws Exception {
      Path pdf = generate(new DocxBuilder().png("rIdP", new byte[] {1, 2, 3, 4})
          .add(p(r("前") + inlinePicture("rIdP", 60, 30) + r("後"))));

      assertThat(PdfLines.of(pdf).get(0).get(0).text()).isEqualTo("前後");
      assertThat(imageCount(pdf)).isZero();
    }

    @Test
    @DisplayName("a drawing which is not a picture (a shape) is skipped, the text is output")
    void shape() throws Exception {
      String shape = "<w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/"
          + "markup-compatibility/2006\"><mc:Choice Requires=\"wps\"><w:drawing/></mc:Choice>"
          + "<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>";
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("前") + shape + r("後"))));

      assertThat(pages.get(0).get(0).text()).isEqualTo("前後");
    }
  }

  @Nested
  @DisplayName("file formats")
  class FileFormats {

    /** Rewrites the content type of the main part of a .docx file (e.g. to a .docm file). */
    private Path withMainContentType(Path docx, String contentType, Path out) throws Exception {
      try (java.util.zip.ZipFile in = new java.util.zip.ZipFile(docx.toFile());
          java.util.zip.ZipOutputStream zip =
              new java.util.zip.ZipOutputStream(Files.newOutputStream(out))) {
        for (var entry : java.util.Collections.list(in.entries())) {
          byte[] data = in.getInputStream(entry).readAllBytes();
          if (entry.getName().equals("[Content_Types].xml")) {
            data = new String(data, java.nio.charset.StandardCharsets.UTF_8).replace(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                contentType).getBytes(java.nio.charset.StandardCharsets.UTF_8);
          }
          zip.putNextEntry(new java.util.zip.ZipEntry(entry.getName()));
          zip.write(data);
          zip.closeEntry();
        }
      }
      return out;
    }

    @org.junit.jupiter.params.ParameterizedTest(name = "{0}")
    @org.junit.jupiter.params.provider.CsvSource({
        "docm, application/vnd.ms-word.document.macroEnabled.main+xml",
        "dotx, application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "dotm, application/vnd.ms-word.template.macroEnabledTemplate.main+xml"})
    @SuppressWarnings("null")
    @DisplayName("macro-enabled documents and templates are converted (macros are not run)")
    void macroEnabledAndTemplates(String extension, String contentType) throws Exception {
      Path docx = new DocxBuilder().add(p(r("本文"))).write(tempDir.resolve("plain.docx"));
      Path word = withMainContentType(docx, contentType, tempDir.resolve("file." + extension));
      Path pdf = tempDir.resolve(extension + ".pdf");
      Path font =
          Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
      WordToPdfUtil.generate(word, pdf, PdfGenerateOptions.builderForExplicitFont(font).build());

      assertThat(PdfLines.of(pdf).get(0).get(0).text()).isEqualTo("本文");
    }
  }

  @Nested
  @DisplayName("files which are not Word files")
  class NotWordFiles {

    @SuppressWarnings("null")
    private void assertNotSupported(Path file) throws Exception {
      Path font =
          Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
      Path pdf = tempDir.resolve("out.pdf");

      assertThatThrownBy(() -> WordToPdfUtil.generate(file, pdf,
          PdfGenerateOptions.builderForExplicitFont(font).build()))
              .isInstanceOfSatisfying(FileFormatNotSupportedException.class,
                  ex -> assertThat(ex.getViolations().getBusinessViolations().get(0)
                      .getMessageArgs()).containsExactly("docx, docm, dotx, dotm"));
      assertThat(pdf).doesNotExist();
    }

    @Test
    @DisplayName("a text file: FileFormatNotSupportedException")
    void text() throws Exception {
      assertNotSupported(Files.writeString(tempDir.resolve("text.docx"), "not a word file"));
    }

    @Test
    @DisplayName("an empty file: FileFormatNotSupportedException")
    void empty() throws Exception {
      assertNotSupported(Files.write(tempDir.resolve("empty.docx"), new byte[0]));
    }

    @Test
    @DisplayName("a legacy .doc file (OLE2 without encryption): FileFormatNotSupportedException")
    void legacyDoc() throws Exception {
      Path doc = tempDir.resolve("legacy.doc");
      try (POIFSFileSystem fs = new POIFSFileSystem();
          OutputStream out = Files.newOutputStream(doc)) {
        fs.createDocument(new java.io.ByteArrayInputStream(new byte[600]), "WordDocument");
        fs.writeFilesystem(out);
      }
      assertNotSupported(doc);
    }

    @Test
    @DisplayName("an Excel file: FileFormatNotSupportedException")
    void excel() throws Exception {
      Path xlsx = tempDir.resolve("book.xlsx");
      try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
          OutputStream out = Files.newOutputStream(xlsx)) {
        wb.createSheet("s");
        wb.write(out);
      }
      assertNotSupported(xlsx);
    }

    @Test
    @DisplayName("a zip file which is not an Office file: FileFormatNotSupportedException")
    void zip() throws Exception {
      Path zip = tempDir.resolve("other.docx");
      try (var out = new java.util.zip.ZipOutputStream(Files.newOutputStream(zip))) {
        out.putNextEntry(new java.util.zip.ZipEntry("a.txt"));
        out.write(1);
        out.closeEntry();
      }
      assertNotSupported(zip);
    }
  }

  @Nested
  @DisplayName("password-protected Word files")
  class Passwords {

    /** Writes the document encrypted with the password. */
    private Path encrypted(String password) throws Exception {
      Path plain = new DocxBuilder().add(p(r("暗号化された文書"))).write(
          tempDir.resolve("plain.docx"));
      Path file = tempDir.resolve("encrypted.docx");
      try (POIFSFileSystem fs = new POIFSFileSystem()) {
        Encryptor encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
        encryptor.confirmPassword(password);
        try (OutputStream out = encryptor.getDataStream(fs);
            InputStream in = Files.newInputStream(plain)) {
          in.transferTo(out);
        }
        try (OutputStream out = Files.newOutputStream(file)) {
          fs.writeFilesystem(out);
        }
      }
      return file;
    }

    @SuppressWarnings("null")
    private PdfGenerateOptions options(@Nullable String wordPassword) throws Exception {
      Path font =
          Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
      return PdfGenerateOptions.builderForExplicitFont(font).wordPassword(wordPassword).build();
    }

    @Test
    @DisplayName("the correct password: the file is converted")
    void correct() throws Exception {
      Path pdf = tempDir.resolve("out.pdf");
      WordToPdfUtil.generate(encrypted("secret"), pdf, options("secret"));

      assertThat(PdfLines.of(pdf).get(0).get(0).text()).isEqualTo("暗号化された文書");
    }

    @Test
    @DisplayName("a wrong password: FilePasswordIncorrectException")
    void wrong() throws Exception {
      Path word = encrypted("secret");

      assertThatThrownBy(() -> WordToPdfUtil.generate(word, tempDir.resolve("out.pdf"),
          options("wrong"))).isInstanceOf(FilePasswordIncorrectException.class);
    }

    @Test
    @DisplayName("no password for a protected file: FilePasswordRequiredException")
    void missing() throws Exception {
      Path word = encrypted("secret");

      assertThatThrownBy(() -> WordToPdfUtil.generate(word, tempDir.resolve("out.pdf"),
          options(null))).isInstanceOf(FilePasswordRequiredException.class);
    }

    @Test
    @DisplayName("a password for a file which is not protected is ignored")
    void notProtected() throws Exception {
      Path word = new DocxBuilder().add(p(r("平文"))).write(tempDir.resolve("plain2.docx"));
      Path pdf = tempDir.resolve("out.pdf");
      WordToPdfUtil.generate(word, pdf, options("secret"));

      assertThat(PdfLines.of(pdf).get(0).get(0).text()).isEqualTo("平文");
    }
  }

  @Nested
  @DisplayName("tracked changes and containers")
  class Containers {

    private String firstLine(String paragraphContent) throws Exception {
      return convert(new DocxBuilder().add("<w:p>" + paragraphContent + "</w:p>")).get(0)
          .get(0).text();
    }

    @Test
    @DisplayName("tracked changes: insertions and moved-to text are output, deletions are not")
    void trackedChanges() throws Exception {
      assertThat(firstLine(r("本文") + "<w:ins w:id=\"1\">" + r("挿入") + "</w:ins>"
          + "<w:del w:id=\"2\"><w:r><w:delText>削除</w:delText></w:r></w:del>"
          + "<w:moveFrom w:id=\"3\">" + r("移動元") + "</w:moveFrom>"
          + "<w:moveTo w:id=\"4\">" + r("移動先") + "</w:moveTo>")).isEqualTo("本文挿入移動先");
    }

    @Test
    @DisplayName("the content of content controls, smart tags and hyperlinks is output")
    void inlineContainers() throws Exception {
      assertThat(firstLine("<w:sdt><w:sdtPr/><w:sdtContent>" + r("制御") + "</w:sdtContent></w:sdt>"
          + "<w:smartTag w:element=\"x\">" + r("タグ") + "</w:smartTag>"
          + "<w:hyperlink w:anchor=\"none\">" + r("リンク") + "</w:hyperlink>"))
              .isEqualTo("制御タグリンク");
    }

    @Test
    @DisplayName("paragraphs and tables in a block content control are output")
    void blockContentControl() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add("<w:sdt><w:sdtPr/><w:sdtContent>"
          + p(r("段落")) + "<w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc>"
          + p(r("表")) + "</w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>").add(p(r("後"))));

      assertThat(pages.get(0)).extracting(Line::text).containsExactly("段落", "表", "後");
    }

    @Test
    @DisplayName("special characters: non-breaking hyphen, soft hyphen, line break, tab in text")
    void specialCharacters() throws Exception {
      List<List<Line>> pages = convert(new DocxBuilder().add(p(r("A")
          + "<w:r><w:noBreakHyphen/></w:r>" + r("B") + "<w:r><w:softHyphen/></w:r>" + r("C")
          + "<w:r><w:cr/></w:r>" + r("改行後\t字"))));

      assertThat(pages.get(0).get(0).text()).isEqualTo("A-BC");
      assertThat(pages.get(0).get(1).leftOf("字")).isCloseTo(LEFT + 36, TOL);
    }

    @Test
    @DisplayName("footnote references and comments are skipped without an error")
    void unsupportedInline() throws Exception {
      assertThat(firstLine(r("本文") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r>"
          + "<w:commentRangeStart w:id=\"1\"/>" + r("注釈付き") + "<w:commentRangeEnd w:id=\"1\"/>"
          + "<w:r><w:commentReference w:id=\"1\"/></w:r>")).isEqualTo("本文注釈付き");
    }
  }

  @Nested
  @DisplayName("table of contents, links and outline")
  class Navigation {

    private static String heading(int level, String bookmark, String text) {
      return p("<w:outlineLvl w:val=\"" + level + "\"/>", "<w:bookmarkStart w:id=\"1\" w:name=\""
          + bookmark + "\"/>" + r(text) + "<w:bookmarkEnd w:id=\"1\"/>");
    }

    private static String tocEntry(String bookmark, String text) {
      return p("<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9000\"/></w:tabs>",
          "<w:hyperlink w:anchor=\"" + bookmark + "\">" + r(text) + "<w:r><w:tab/></w:r>"
              + field("PAGEREF " + bookmark + " \\h", "99") + "</w:hyperlink>");
    }

    private DocxBuilder documentWithToc() {
      return new DocxBuilder().add(tocEntry("_h1", "第一章")).add(tocEntry("_h2", "第一節"))
          .add(tocEntry("_h3", "第二章")).add(p(pageBreak())).add(heading(0, "_h1", "第一章"))
          .add(heading(1, "_h2", "第一節")).add(p(pageBreak())).add(p(pageBreak()))
          .add(heading(0, "_h3", "第二章"));
    }

    @Test
    @DisplayName("each entry shows the page of its heading, right-aligned after the leader")
    void pageNumbers() throws Exception {
      List<Line> toc = convert(documentWithToc()).get(0);

      assertThat(toc.get(0).text()).startsWith("第一章.....").endsWith("2");
      assertThat(toc.get(1).text()).startsWith("第一節.....").endsWith("2");
      assertThat(toc.get(2).text()).startsWith("第二章.....").endsWith("4");
      assertThat(toc.get(2).right()).isCloseTo(LEFT + 450, TOL);
    }

    @Test
    @DisplayName("each entry links to the page of its heading")
    void links() throws Exception {
      Path pdf = generate(documentWithToc());

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        List<Integer> targets = new ArrayList<>();
        for (PDAnnotation a : doc.getPage(0).getAnnotations()) {
          if (a instanceof PDAnnotationLink link
              && link.getAction() instanceof org.apache.pdfbox.pdmodel.interactive.action
                  .PDActionGoTo action
              && action.getDestination() instanceof org.apache.pdfbox.pdmodel.interactive
                  .documentnavigation.destination.PDPageDestination dest) {
            targets.add(doc.getPages().indexOf(dest.getPage()));
          }
        }
        assertThat(targets).contains(1, 3).doesNotContain(0, 2);
      }
    }

    @Test
    @DisplayName("headings make the outline of the PDF, nested by their levels")
    void outline() throws Exception {
      Path pdf = generate(documentWithToc());

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        PDOutlineItem first = doc.getDocumentCatalog().getDocumentOutline().getFirstChild();
        assertThat(first.getTitle()).isEqualTo("第一章");
        assertThat(first.getFirstChild().getTitle()).isEqualTo("第一節");
        assertThat(doc.getPages().indexOf(first.findDestinationPage(doc))).isEqualTo(1);
        PDOutlineItem second = first.getNextSibling();
        assertThat(second.getTitle()).isEqualTo("第二章");
        assertThat(doc.getPages().indexOf(second.findDestinationPage(doc))).isEqualTo(3);
        assertThat(second.getNextSibling()).isNull();
      }
    }

    @Test
    @DisplayName("the outline level of a paragraph style makes headings, with their numbers")
    void outlineFromStyle() throws Exception {
      String styles = "<w:style w:type=\"paragraph\" w:styleId=\"H\"><w:name w:val=\"heading 1\"/>"
          + "<w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr><w:outlineLvl w:val=\"0\"/></w:pPr>"
          + "</w:style>";
      String numbering = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">"
          + "<w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:suff w:val=\"space\"/>"
          + "<w:lvlText w:val=\"第%1条\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\">"
          + "<w:abstractNumId w:val=\"0\"/></w:num>";
      Path pdf = generate(new DocxBuilder().styles(styles).numbering(numbering)
          .add(p("<w:pStyle w:val=\"H\"/>", r("目的"))).add(p(r("本文")))
          .add(p("<w:pStyle w:val=\"H\"/>", r("定義"))));

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        PDOutlineItem first = doc.getDocumentCatalog().getDocumentOutline().getFirstChild();
        assertThat(first.getTitle()).isEqualTo("第1条 目的");
        assertThat(first.getNextSibling().getTitle()).isEqualTo("第2条 定義");
      }
    }

    @Test
    @DisplayName("a document without headings has no outline")
    void noOutline() throws Exception {
      Path pdf = generate(new DocxBuilder().add(p(r("本文"))));

      try (PDDocument doc = Loader.loadPDF(pdf.toFile())) {
        assertThat(doc.getDocumentCatalog().getDocumentOutline()).isNull();
      }
    }

    @Test
    @DisplayName("a reference to a bookmark which does not exist shows the stored result")
    void missingBookmark() throws Exception {
      List<Line> page = convert(new DocxBuilder().add(tocEntry("_none", "見出し"))).get(0);

      assertThat(page.get(0).text()).endsWith("99");
    }
  }
}
