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
 * Thrown when the number of header columns found in the Excel table differs from the number
 * of expected header labels.
 */
public class NumberOfTableHeadersDifferException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name
   * @param actualColumnCount the number of header columns actually found in the Excel table
   * @param expectedColumnCount the number of header labels expected
   */
  public NumberOfTableHeadersDifferException(String sheetName, int actualColumnCount,
      int expectedColumnCount) {
    super("jp.ecuacion.util.excel.NumberOfTableHeadersDiffer.message", sheetName,
        Integer.toString(actualColumnCount), Integer.toString(expectedColumnCount));
  }
}
