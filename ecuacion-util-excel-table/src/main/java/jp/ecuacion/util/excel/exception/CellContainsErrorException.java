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

import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import org.apache.commons.lang3.StringUtils;
import org.jspecify.annotations.Nullable;

/**
 * Thrown when a cell contains an error value (e.g. {@code #NUM!}, {@code #DIV/0!}).
 */
public class CellContainsErrorException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name
   * @param cellAddress the address of the cell that contains an error, e.g. {@code "A1"}
   * @param filename filename or file path of the Excel file to add to the message, or
   *     {@code null} when unavailable
   */
  public CellContainsErrorException(String sheetName, String cellAddress,
      @Nullable String filename) {
    super("jp.ecuacion.util.excel.CellContainsError.message", sheetName, cellAddress,
        StringUtils.isEmpty(filename) ? "" : messageItemSeparator(),
        StringUtils.isEmpty(filename) ? "" : Arg.message("jp.ecuacion.util.excel.common.filename",
            filename));
  }

  private static Arg messageItemSeparator() {
    return Arg.message("jp.ecuacion.util.excel.common.messageItemSeparator");
  }
}
