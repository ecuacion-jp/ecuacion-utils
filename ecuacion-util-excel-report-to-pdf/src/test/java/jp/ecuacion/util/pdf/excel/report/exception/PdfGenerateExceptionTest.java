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
package jp.ecuacion.util.pdf.excel.report.exception;

import static org.assertj.core.api.Assertions.assertThat;
import java.nio.file.Path;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

@DisplayName("PdfGenerateException subclasses")
public class PdfGenerateExceptionTest {

  @Test
  @DisplayName("SheetNotExistException stores messageId and messageArgs")
  void sheetNotExist() {
    SheetNotExistException ex = new SheetNotExistException("invoice");

    assertThat(ex.getMessageId())
        .isEqualTo("jp.ecuacion.util.pdf.excel.report.SheetNotExist.message");
    assertThat(ex.getViolations().getBusinessViolations().get(0).getMessageArgs())
        .containsExactly("invoice");
  }

  @Test
  @DisplayName("SheetHasNoPrintAreaException stores messageId and messageArgs")
  void sheetHasNoPrintArea() {
    SheetHasNoPrintAreaException ex = new SheetHasNoPrintAreaException("invoice");

    assertThat(ex.getMessageId())
        .isEqualTo("jp.ecuacion.util.pdf.excel.report.SheetHasNoPrintArea.message");
    assertThat(ex.getViolations().getBusinessViolations().get(0).getMessageArgs())
        .containsExactly("invoice");
  }

  @Test
  @DisplayName("SystemFontNotFoundException stores messageId and messageArgs")
  void systemFontNotFound() {
    SystemFontNotFoundException ex = new SystemFontNotFoundException("Calibri");

    assertThat(ex.getMessageId())
        .isEqualTo("jp.ecuacion.util.pdf.excel.report.SystemFontNotFound.message");
    assertThat(ex.getViolations().getBusinessViolations().get(0).getMessageArgs())
        .containsExactly("Calibri");
  }

  @Test
  @DisplayName("FontLoadFailedException stores messageId and messageArgs")
  void fontLoadFailed() {
    Path fontFile = Path.of("/fonts/Calibri.ttf");
    FontLoadFailedException ex = new FontLoadFailedException("Calibri", fontFile);

    assertThat(ex.getMessageId())
        .isEqualTo("jp.ecuacion.util.pdf.excel.report.FontLoadFailed.message");
    assertThat(ex.getViolations().getBusinessViolations().get(0).getMessageArgs())
        .containsExactly("Calibri", fontFile);
  }

  @Test
  @DisplayName("CharacterNotRenderableException stores messageId and messageArgs")
  void characterNotRenderable() {
    CharacterNotRenderableException ex =
        new CharacterNotRenderableException("E000", "", "Calibri", "(none configured)");

    assertThat(ex.getMessageId())
        .isEqualTo("jp.ecuacion.util.pdf.excel.report.CharacterNotRenderable.message");
    assertThat(ex.getViolations().getBusinessViolations().get(0).getMessageArgs())
        .containsExactly("E000", "", "Calibri", "(none configured)");
  }
}
