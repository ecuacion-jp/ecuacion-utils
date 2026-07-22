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
 * Thrown when {@code useSystemFonts} is enabled and the workbook's default font is not installed
 * on the running system, with no fallback font configured.
 */
public class SystemFontNotFoundException extends PdfGenerateException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param fontName the font name that could not be found on the system
   */
  public SystemFontNotFoundException(String fontName) {
    super("jp.ecuacion.util.pdf.excel.report.SystemFontNotFound.message", fontName);
  }
}
