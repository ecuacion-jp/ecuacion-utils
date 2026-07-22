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
 * Thrown when a sheet has neither a defined print area nor any data to infer one from.
 */
public class SheetHasNoPrintAreaException extends PdfGenerateException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name that has no print area and no data
   */
  public SheetHasNoPrintAreaException(String sheetName) {
    super("jp.ecuacion.util.pdf.excel.report.SheetHasNoPrintArea.message", sheetName);
  }
}
