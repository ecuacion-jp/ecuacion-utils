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
package jp.ecuacion.util.pdf.excel.report.options;

import static org.assertj.core.api.Assertions.assertThat;
import java.net.URISyntaxException;
import java.nio.file.Path;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("PdfGenerateOptions")
public class PdfGenerateOptionsTest {

  private static Path regularFontPath() throws URISyntaxException {
    return Path.of(PdfGenerateOptionsTest.class
        .getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf").toURI());
  }

  @Nested
  @DisplayName("Builder: password setters")
  class PasswordSetters {

    @Test
    @DisplayName("excelPassword() is stored and returned by getExcelPassword()")
    void excelPassword() throws Exception {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .excelPassword("excel-pass")
          .build();

      assertThat(opts.getExcelPassword()).isEqualTo("excel-pass");
    }

    @Test
    @DisplayName("pdfPassword() is stored and returned by getPdfPassword()")
    void pdfPassword() throws Exception {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .pdfPassword("pdf-pass")
          .build();

      assertThat(opts.getPdfPassword()).isEqualTo("pdf-pass");
    }

    @Test
    @DisplayName("pdfOwnerPassword() is stored and returned by getPdfOwnerPassword()")
    void pdfOwnerPassword() throws Exception {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .pdfOwnerPassword("owner-pass")
          .build();

      assertThat(opts.getPdfOwnerPassword()).isEqualTo("owner-pass");
    }

    @Test
    @DisplayName("password not set → getter returns null")
    void passwordNotSet() throws Exception {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .build();

      assertThat(opts.getExcelPassword()).isNull();
      assertThat(opts.getPdfPassword()).isNull();
      assertThat(opts.getPdfOwnerPassword()).isNull();
    }
  }

  @Nested
  @DisplayName("Builder: addRegularFontPath / addBoldFontPath")
  class RegularAndBoldFontPaths {

    @Test
    @DisplayName("builderForExplicitFont — the given path is the first regularFontPaths entry")
    void explicitFontIsFirstRegularEntry() throws Exception {
      Path regular = regularFontPath();
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regular).build();

      assertThat(opts.getRegularFontPaths()).containsExactly(regular);
      assertThat(opts.getBoldFontPaths()).isEmpty();
    }

    @Test
    @DisplayName("addRegularFontPath/addBoldFontPath called twice — returns entries in order")
    void calledTwiceReturnsEntriesInOrder() throws Exception {
      Path first = regularFontPath();
      Path second = regularFontPath();
      Path bold = regularFontPath();
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(first)
          .addRegularFontPath(second)
          .addBoldFontPath(bold)
          .build();

      assertThat(opts.getRegularFontPaths()).containsExactly(first, second);
      assertThat(opts.getBoldFontPaths()).containsExactly(bold);
    }
  }

  @Nested
  @DisplayName("Builder.build()")
  class BuildValidation {

    @Test
    @DisplayName("builderForSystemFonts().build() with no regularFontPath → succeeds")
    void buildWithUseSystemFontsNoFontPath() {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForSystemFonts().build();

      assertThat(opts.isUseSystemFonts()).isTrue();
      assertThat(opts.getRegularFontPaths()).isEmpty();
    }
  }
}
