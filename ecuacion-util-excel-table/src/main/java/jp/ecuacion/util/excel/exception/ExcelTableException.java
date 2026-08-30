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
import jp.ecuacion.lib.core.util.PropertiesFileUtil;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.Violations;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.jspecify.annotations.Nullable;

/**
 * Is the common superclass of exceptions thrown when a table-related error occurs in
 * {@code ecuacion-util-excel-table}.
 *
 * <p>Each specific failure is represented by one of the concrete subclasses permitted below
 * (e.g. {@link SheetNotExistException}, {@link CellContainsErrorException}), so callers can
 * {@code catch} the specific case they want to handle differently instead of branching on a
 * {@code messageId} string. Catching this class itself still works for callers that only want
 * to handle "some table error" generically. Being {@code sealed}, a {@code switch} over all
 * permitted subclasses is exhaustive without a {@code default} branch, so the compiler flags
 * any case left unhandled when a new subclass is added.</p>
 */
public abstract sealed class ExcelTableException extends ViolationException
    permits ColumnSizeIsZeroException, ExcelFeatureNotImplementedException,
    ExternalWorkbookNotFoundException, FarLeftHeaderLabelNotFoundException,
    FormulaEvaluationUnknownErrorException, HeaderCellIsBlankException,
    NumberOfTableHeadersDifferException, SheetNotExistException, TableHeaderTitleWrongException,
    CellContainsErrorException {

  private static final long serialVersionUID = 1L;

  private static final String CELL_FORMAT_R1C1_KEY = "jp.ecuacion.util.excel.cell-format-r1c1";

  /**
   * Constructs an instance.
   *
   * @param messageId messageId
   * @param messageArgs messageArgs
   */
  protected ExcelTableException(String messageId, @Nullable Object... messageArgs) {
    super(new Violations().add(messageId, messageArgs));
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
   * Builds a single {@link Arg} representing a cell position, in the same way as
   * {@link #cellPositionArg(int, int)}, deriving the row and column from {@code cell}.
   *
   * @param cell the cell to build the position from
   * @return an {@code Arg} embeddable as a single {@code {n}} placeholder in a message template
   */
  protected static Arg cellPositionArg(Cell cell) {
    return cellPositionArg(cell.getRowIndex() + 1, cell.getColumnIndex() + 1);
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
