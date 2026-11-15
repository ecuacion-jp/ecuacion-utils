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
import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.common.exception.FileFormatNotSupportedException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordIncorrectException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordRequiredException;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Tests the exceptions for the failures caused by the Excel file itself: its format and its
 * password.
 */
@DisplayName("ExcelToPdfUtil: files which cannot be opened")
class SourceFileErrorTest {

  @TempDir
  Path tempDir;

  private Path generate(Path excel, @Nullable String password) throws Exception {
    Path font = Path.of(getClass().getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
    Path pdf = tempDir.resolve("out.pdf");
    ExcelToPdfUtil.generate(excel, List.of("s"), pdf,
        PdfGenerateOptions.builderForExplicitFont(font).excelPassword(password).build());
    return pdf;
  }

  private Path workbook(String name) throws Exception {
    Path file = tempDir.resolve(name);
    try (XSSFWorkbook wb = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file)) {
      wb.createSheet("s").createRow(0).createCell(0).setCellValue("本文");
      wb.write(out);
    }
    return file;
  }

  @Nested
  @DisplayName("files which are not Excel files")
  class NotExcelFiles {

    private void assertNotSupported(Path file) {
      assertThatThrownBy(() -> generate(file, null))
          .isInstanceOfSatisfying(FileFormatNotSupportedException.class,
              ex -> assertThat(ex.getViolations().getBusinessViolations().get(0)
                  .getMessageArgs()).containsExactly("xlsx, xlsm, xls"));
    }

    @Test
    @DisplayName("a text file")
    void text() throws Exception {
      assertNotSupported(Files.writeString(tempDir.resolve("text.xlsx"), "not an excel file"));
    }

    @Test
    @DisplayName("an empty file")
    void empty() throws Exception {
      assertNotSupported(Files.write(tempDir.resolve("empty.xlsx"), new byte[0]));
    }

    @Test
    @DisplayName("a legacy Word (.doc) file")
    void legacyWord() throws Exception {
      Path doc = tempDir.resolve("legacy.doc");
      try (POIFSFileSystem fs = new POIFSFileSystem();
          OutputStream out = Files.newOutputStream(doc)) {
        fs.createDocument(new ByteArrayInputStream(new byte[600]), "WordDocument");
        fs.writeFilesystem(out);
      }
      assertNotSupported(doc);
    }

    @Test
    @DisplayName("a zip file which is not an Office file")
    void zip() throws Exception {
      Path zip = tempDir.resolve("other.xlsx");
      try (var out = new java.util.zip.ZipOutputStream(Files.newOutputStream(zip))) {
        out.putNextEntry(new java.util.zip.ZipEntry("a.txt"));
        out.write(1);
        out.closeEntry();
      }
      assertNotSupported(zip);
    }

    @Test
    @DisplayName("a missing file is not a format problem: UncheckedIOException")
    void missing() {
      assertThatThrownBy(() -> generate(tempDir.resolve("missing.xlsx"), null))
          .isInstanceOf(UncheckedIOException.class);
    }
  }

  @Nested
  @DisplayName("Excel formats which are supported")
  class SupportedFormats {

    @Test
    @DisplayName("a legacy .xls file is converted")
    void xls() throws Exception {
      Path xls = tempDir.resolve("old.xls");
      try (HSSFWorkbook wb = new HSSFWorkbook(); OutputStream out = Files.newOutputStream(xls)) {
        wb.createSheet("s").createRow(0).createCell(0).setCellValue("本文");
        wb.write(out);
      }
      assertThat(generate(xls, null)).exists();
    }
  }

  @Nested
  @DisplayName("password-protected Excel files")
  class Passwords {

    private Path encrypted(String password) throws Exception {
      Path plain = workbook("plain.xlsx");
      Path file = tempDir.resolve("encrypted.xlsx");
      try (POIFSFileSystem fs = new POIFSFileSystem()) {
        Encryptor encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
        encryptor.confirmPassword(password);
        try (OutputStream out = encryptor.getDataStream(fs)) {
          Files.copy(plain, out);
        }
        try (OutputStream out = Files.newOutputStream(file)) {
          fs.writeFilesystem(out);
        }
      }
      return file;
    }

    @Test
    @DisplayName("the correct password: the file is converted")
    void correct() throws Exception {
      assertThat(generate(encrypted("secret"), "secret")).exists();
    }

    @Test
    @DisplayName("no password: FilePasswordRequiredException")
    void missing() throws Exception {
      Path file = encrypted("secret");

      assertThatThrownBy(() -> generate(file, null))
          .isInstanceOf(FilePasswordRequiredException.class);
    }

    @Test
    @DisplayName("a wrong password: FilePasswordIncorrectException")
    void wrong() throws Exception {
      Path file = encrypted("secret");

      assertThatThrownBy(() -> generate(file, "wrong"))
          .isInstanceOf(FilePasswordIncorrectException.class);
    }

    @Test
    @DisplayName("a password for a file which is not protected is ignored")
    void notProtected() throws Exception {
      assertThat(generate(workbook("plain2.xlsx"), "secret")).exists();
    }
  }
}
