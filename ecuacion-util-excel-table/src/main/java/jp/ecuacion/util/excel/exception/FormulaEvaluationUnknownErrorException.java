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
package jp.ecuacion.util.excel.exception;

/**
 * Thrown when an unrecognized error occurs while evaluating a formula, wrapping whatever
 * exception the underlying Excel manipulation library (Apache POI) raised.
 */
public class FormulaEvaluationUnknownErrorException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param fileInfoArg filename or file path of the Excel file being evaluated, or a localized
   *     "(none)" label when unavailable
   * @param sheetName the sheet name
   * @param cellAddress the address of the cell containing the formula
   * @param detail a newline-joined dump of the underlying exception's message chain
   */
  public FormulaEvaluationUnknownErrorException(Object fileInfoArg, String sheetName,
      String cellAddress, String detail) {
    super("jp.ecuacion.util.excel.ExcelWriteUtil.DetailUnknown.message", fileInfoArg, sheetName,
        cellAddress, detail);
  }
}
