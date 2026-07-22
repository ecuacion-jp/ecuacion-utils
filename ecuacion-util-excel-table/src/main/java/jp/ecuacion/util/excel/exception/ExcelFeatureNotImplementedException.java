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
 * Thrown when the underlying Excel manipulation library (Apache POI) does not support a
 * feature used by a formula being evaluated.
 */
public class ExcelFeatureNotImplementedException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name
   * @param cellAddress the address of the cell containing the unsupported formula
   * @param reason a description of the unsupported feature (e.g. an unimplemented function name),
   *     or a localized "(unknown)" label when the reason cannot be determined
   * @param fileInfoArg filename or file path of the Excel file being evaluated, or a localized
   *     "(none)" label when unavailable
   */
  public ExcelFeatureNotImplementedException(String sheetName, String cellAddress, Object reason,
      Object fileInfoArg) {
    super("jp.ecuacion.util.excel.ExcelWriteUtil.NotImplementedException.message", sheetName,
        cellAddress, reason, fileInfoArg);
  }
}
