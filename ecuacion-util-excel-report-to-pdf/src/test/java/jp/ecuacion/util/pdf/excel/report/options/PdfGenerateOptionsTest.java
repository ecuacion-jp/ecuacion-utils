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
  @DisplayName("Builder: addFallbackFont")
  class AdditionalFallbackFonts {

    @Test
    @DisplayName("not called — getAdditionalFallbackFonts() returns an empty list")
    void notCalledReturnsEmptyList() throws Exception {
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .build();

      assertThat(opts.getAdditionalFallbackFonts()).isEmpty();
    }

    @Test
    @DisplayName("called twice — returns both entries in registration order")
    void calledTwiceReturnsEntriesInOrder() throws Exception {
      Path first = regularFontPath();
      Path second = regularFontPath();
      PdfGenerateOptions opts = PdfGenerateOptions.builderForExplicitFont(regularFontPath())
          .addFallbackFont(first, null)
          .addFallbackFont(second, first)
          .build();

      assertThat(opts.getAdditionalFallbackFonts()).hasSize(2);
      assertThat(opts.getAdditionalFallbackFonts().get(0).regularFontPath()).isEqualTo(first);
      assertThat(opts.getAdditionalFallbackFonts().get(0).boldFontPath()).isNull();
      assertThat(opts.getAdditionalFallbackFonts().get(1).regularFontPath()).isEqualTo(second);
      assertThat(opts.getAdditionalFallbackFonts().get(1).boldFontPath()).isEqualTo(first);
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
      assertThat(opts.getRegularFontPath()).isNull();
    }
  }
}
