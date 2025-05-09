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
package jp.ecuacion.util.poi.excel.exception;

import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.util.PropertyFileUtil.Arg;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Provides {@code BizLogicAppException} with location and exception of cause.
 */
public class ExcelAppException extends BizLogicAppException {

  private static final long serialVersionUID = 1L;

  private Workbook workbook;
  private Sheet sheet;
  private Cell cell;

  /**
   * Constructs an instance.
   */
  public ExcelAppException(String messageId, String... messageArgs) {
    super(messageId, messageArgs);
  }

  /**
   * Constructs an instance.
   */
  public ExcelAppException(String messageId, Arg... messageArgs) {
    super(messageId, messageArgs);
  }

  public Workbook getWorkbook() {
    return workbook;
  }

  /**
   * Sets workdbook and returns ExcelAppException for method chain.
   * 
   * @param workbook workbook to set.
   * @return ExcelAppException
   */
  public ExcelAppException workbook(Workbook workbook) {
    this.workbook = workbook;

    return this;
  }

  public Sheet getSheet() {
    return sheet;
  }

  /**
   * Sets workdbook and returns ExcelAppException for method chain.
   * 
   * @param sheet sheet to set.
   * @return ExcelAppException
   */
  public ExcelAppException sheet(Sheet sheet) {
    this.sheet = sheet;
    this.workbook = sheet.getWorkbook();
    
    return this;
  }

  public Cell getCell() {
    return cell;
  }

  /**
   * Sets workdbook and returns ExcelAppException for method chain.
   * 
   * @param cell cell to set.
   * @return ExcelAppException
   */
  public ExcelAppException cell(Cell cell) {
    this.cell = cell;
    this.sheet = cell.getSheet();
    this.workbook = sheet.getWorkbook();
    
    return this;
  }

  /**
   * Sets cause exception and return self for method chain.
   */
  public ExcelAppException cause(Throwable th) {
    return (ExcelAppException) super.cause(th);
  }
}
