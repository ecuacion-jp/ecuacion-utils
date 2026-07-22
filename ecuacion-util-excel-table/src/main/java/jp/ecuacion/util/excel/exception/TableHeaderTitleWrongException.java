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

import org.jspecify.annotations.Nullable;

/**
 * Thrown when a header cell's label in the Excel table differs from the expected label at
 * that position.
 */
public class TableHeaderTitleWrongException extends ExcelTableException {

  private static final long serialVersionUID = 1L;

  /**
   * Constructs an instance.
   *
   * @param sheetName the sheet name
   * @param position the 1-based column position of the header, from the user's perspective
   * @param actualLabel the header label actually found in the Excel table, may be {@code null}
   *     when the cell is blank
   * @param expectedLabel the header label expected at that position
   */
  public TableHeaderTitleWrongException(String sheetName, int position,
      @Nullable String actualLabel, String expectedLabel) {
    super("jp.ecuacion.util.excel.TableHeaderTitleWrong.message", sheetName,
        Integer.toString(position), actualLabel, expectedLabel);
  }
}
