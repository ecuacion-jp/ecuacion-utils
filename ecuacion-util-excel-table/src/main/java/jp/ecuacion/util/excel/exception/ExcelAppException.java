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

import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.BusinessViolation;
import jp.ecuacion.lib.core.violation.Violations;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Provides {@code ViolationException} for Excel errors with workbook, sheet, and cell context.
 */
public class ExcelAppException extends ViolationException {

  private static final long serialVersionUID = 1L;

  private @Nullable Workbook workbook;
  private @Nullable Sheet sheet;
  private @Nullable Cell cell;

  /**
   * Constructs an instance.
   *
   * @param messageId messageId
   * @param messageArgs messageArgs
   */
  public ExcelAppException(String messageId, @Nullable String... messageArgs) {
    super(new Violations().add(new BusinessViolation(messageId, messageArgs)));
  }

  /**
   * Constructs an instance.
   *
   * @param messageId messageId
   * @param messageArgs messageArgs
   */
  public ExcelAppException(String messageId, Arg... messageArgs) {
    super(new Violations().add(new BusinessViolation(messageId, messageArgs)));
  }

  /**
   * Gets messageId.
   *
   * @return messageId
   */
  public String getMessageId() {
    return getViolations().getBusinessViolations().get(0).getMessageId();
  }

  /**
   * Gets workbook.
   *
   * @return workbook
   */
  public @Nullable Workbook getWorkbook() {
    return workbook;
  }

  /**
   * Sets workbook and returns ExcelAppException for method chain.
   *
   * @param workbook workbook to set.
   * @return ExcelAppException
   */
  public ExcelAppException workbook(Workbook workbook) {
    this.workbook = workbook;

    return this;
  }

  /**
   * Gets sheet.
   *
   * @return sheet
   */
  public @Nullable Sheet getSheet() {
    return sheet;
  }

  /**
   * Sets sheet and returns ExcelAppException for method chain.
   *
   * @param sheet sheet to set.
   * @return ExcelAppException
   */
  public ExcelAppException sheet(Sheet sheet) {
    this.sheet = sheet;
    this.workbook = sheet.getWorkbook();

    return this;
  }

  /**
   * Gets cell.
   *
   * @return cell
   */
  public @Nullable Cell getCell() {
    return cell;
  }

  /**
   * Sets cell and returns ExcelAppException for method chain.
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
   * Sets cause exception and returns self for method chain.
   *
   * @param th throwable to set as cause
   * @return ExcelAppException
   */
  public ExcelAppException cause(Throwable th) {
    initCause(th);
    return this;
  }
}
