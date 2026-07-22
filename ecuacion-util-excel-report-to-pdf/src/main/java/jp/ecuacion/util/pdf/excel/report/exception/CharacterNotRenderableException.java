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

/**
 * Thrown when a character in a cell cannot be encoded by the primary font or any configured
 * fallback font.
 */
public class CharacterNotRenderableException extends PdfGenerateException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param codePointHex the unrenderable Unicode code point, formatted as an uppercase hex string
   * @param character the unrenderable character itself
   * @param primaryFontDescription a description of the primary font that was tried
   * @param fallbackFontsDescription a description of the fallback fonts that were tried
   */
  public CharacterNotRenderableException(String codePointHex, String character,
      String primaryFontDescription, String fallbackFontsDescription) {
    super("jp.ecuacion.util.pdf.excel.report.CharacterNotRenderable.message", codePointHex,
        character, primaryFontDescription, fallbackFontsDescription);
  }
}
