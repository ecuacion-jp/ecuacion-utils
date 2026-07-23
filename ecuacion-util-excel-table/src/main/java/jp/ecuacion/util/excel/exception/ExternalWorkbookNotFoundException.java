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
 * Thrown when a formula references an external Excel file that cannot be found while
 * evaluating it.
 */
public class ExternalWorkbookNotFoundException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name
   * @param cellAddress the address of the cell containing the formula
   * @param formula the formula referencing the missing external workbook
   * @param fileInfoArg filename or file path of the Excel file being evaluated, or a localized
   *     "(none)" label when unavailable
   */
  public ExternalWorkbookNotFoundException(String sheetName, String cellAddress, String formula,
      Object fileInfoArg) {
    super("jp.ecuacion.util.excel.ExcelWriteUtil.WorkbookNotFoundException.message", sheetName,
        cellAddress, formula, fileInfoArg);
  }
}
