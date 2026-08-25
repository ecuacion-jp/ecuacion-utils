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

import java.util.Objects;
import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.lib.core.util.PropertiesFileUtil;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.BusinessViolation;
import jp.ecuacion.lib.core.violation.Violations;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.jspecify.annotations.Nullable;

/**
 * Is the common superclass of exceptions thrown when a table-related error occurs in
 * {@code ecuacion-util-excel-table}.
 *
 * <p>Each specific failure is represented by one of the concrete subclasses in this package
 * (e.g. {@link SheetNotExistException}, {@link CellContainsErrorException}), so callers can
 * {@code catch} the specific case they want to handle differently instead of branching on a
 * {@code messageId} string. Catching this class itself still works for callers that only want
 * to handle "some table error" generically.</p>
 */
public abstract class ExcelTableException extends ViolationException {

  private static final long serialVersionUID = 1L;

  private static final String CELL_FORMAT_R1C1_KEY = "jp.ecuacion.util.excel.cell-format-r1c1";

  private @Nullable Workbook workbook;
  private @Nullable Sheet sheet;
  private @Nullable Cell cell;

  /**
   * Constructs an instance.
   *
   * @param messageId messageId
   * @param messageArgs messageArgs
   */
  protected ExcelTableException(String messageId, @Nullable Object... messageArgs) {
    super(new Violations().add(new BusinessViolation(messageId, messageArgs)));
  }

  /**
   * Builds a single {@link Arg} representing a cell position, resolved either as an A1-style
   * address (e.g. {@code "B3"}) or as row/column numbers, depending on the
   * {@code jp.ecuacion.util.excel.cell-format-r1c1} application property ({@code "true"} = row
   * and column numbers, default {@code "false"} = A1 address).
   *
   * @param row the 1-based Excel row of the cell
   * @param column the 1-based Excel column of the cell
   * @return an {@code Arg} embeddable as a single {@code {n}} placeholder in a message template
   */
  protected static Arg cellPositionArg(int row, int column) {
    boolean r1c1 =
        Boolean.parseBoolean(PropertiesFileUtil.getApplicationOrElse(CELL_FORMAT_R1C1_KEY,
            "false"));
    String msgId = r1c1 ? "jp.ecuacion.util.excel.common.cellPositionR1c1.message"
        : "jp.ecuacion.util.excel.common.cellPositionA1.message";
    String a1 = new CellAddress(row - 1, column - 1).formatAsString();

    // Integer.toString(...) here, not raw int, since java.text.MessageFormat would otherwise
    // render an int with locale-dependent grouping separators (same pattern as ADR 0003).
    return Arg.message(msgId, a1, Integer.toString(row), Integer.toString(column));
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
   * Sets workbook and returns self for method chain.
   *
   * @param workbook workbook to set.
   * @return ExcelTableException
   */
  public ExcelTableException workbook(Workbook workbook) {
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
   * Sets sheet and returns self for method chain.
   *
   * @param sheet sheet to set.
   * @return ExcelTableException
   */
  public ExcelTableException sheet(Sheet sheet) {
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
   * Sets cell and returns self for method chain.
   *
   * @param cell cell to set.
   * @return ExcelTableException
   */
  public ExcelTableException cell(Cell cell) {
    this.cell = cell;
    this.sheet = cell.getSheet();
    this.workbook = Objects.requireNonNull(sheet).getWorkbook();
    return this;
  }

  /**
   * Sets cause exception and returns self for method chain.
   *
   * @param th throwable to set as cause
   * @return ExcelTableException
   */
  public ExcelTableException cause(Throwable th) {
    initCause(th);
    return this;
  }
}
