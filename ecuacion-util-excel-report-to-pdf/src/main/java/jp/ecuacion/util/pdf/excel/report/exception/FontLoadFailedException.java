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

import java.nio.file.Path;

/**
 * Thrown when a system font file was located but could not be loaded as a
 * {@code TrueTypeFont}.
 */
public class FontLoadFailedException extends PdfGenerateException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param fontName the font name that failed to load
   * @param fontFile the font file that failed to load
   */
  public FontLoadFailedException(String fontName, Path fontFile) {
    super("jp.ecuacion.util.pdf.excel.report.FontLoadFailed.message", fontName, fontFile);
  }
}
