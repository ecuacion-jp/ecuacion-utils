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
package jp.ecuacion.util.poi.excel.util;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.io.File;
import java.io.IOException;
import java.text.Format;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.exception.LoopBreakException;
import jp.ecuacion.util.poi.excel.table.ExcelTable.ContextContainer;
import jp.ecuacion.util.poi.excel.table.reader.ExcelTableReader;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Provides excel reading related {@code apache POI} utility methods.
 */
public class ExcelReadUtil {

  private boolean suppressesWarnLog = false;

  private DetailLogger detailLog = new DetailLogger(this);

  /* 空文字オブジェクトが複数作られると非効率なので定義しておく。 */
  private static final String EMPTY_STRING = "";

  private String noDataString;

  private DateTimeFormatter defaultDateTimeFormat = DateTimeFormatter.ofPattern("yyyy-MM-dd");

  /**
   * Constructs a new instance with {@code NoDataString = NULL}.
   * 
   * <p>{@code NoDataString = NULL} is recommended. 
   *     See {@link jp.ecuacion.util.poi.excel.enums.NoDataString}.</p>
   */
  public ExcelReadUtil() {
    this.noDataString = null;
  }

  /**
   * Constructs a new instance with designated {@code NoDataString}.
   * 
   * <p>{@code NoDataString = NULL} is recommended. 
   *     See {@link jp.ecuacion.util.poi.excel.enums.NoDataString}.</p>
   *     
   * @param noDataString noDataString
   */
  public ExcelReadUtil(@Nonnull NoDataString noDataString) {
    ObjectsUtil.paramRequireNonNull(noDataString);

    this.noDataString = (noDataString == NoDataString.EMPTY_STRING) ? EMPTY_STRING : null;
  }

  /**
   * Sets defaultDateTimeFormat.
   * 
   * @param dateTimeFormat dateTimeFormat
   */
  public void setDefaultDateTimeFormat(DateTimeFormatter dateTimeFormat) {
    this.defaultDateTimeFormat = dateTimeFormat;
  }

  /** 
   * Returns proper {@code NoDataString} value if the argument value is 
   *     {@code null} or {@code ""}, otherwize returns the argument value.
   * 
   * @param value value, may be {@code null}.
   * @return the argument value or the empty string designated by {@code noDataString} 
   *     when the argument value is empty.
   */
  public @Nullable String getNoDataStringIfNoData(@Nullable String value) {
    if (value == null || value.equals("")) {
      return noDataString;

    } else {
      return value;
    }
  }

  /**
  * Returns {@code String} format cell value
  * in spite of the format or value kind of the cell.
  *
  * @param cell the cell of the excel file
  * @return the string which expresses the value of the cell.
   * @throws BizLogicAppException BizLogicAppException
  */
  public @Nullable String getStringFromCell(@Nullable Cell cell) throws BizLogicAppException {
    return getStringFromCell(cell, defaultDateTimeFormat);
  }

  /**
   * Returns {@code String} format cell value 
   *     in spite of the format or value kind of the cell.
   * 
   * @param cell the cell of the excel file
   * @param dateTimeFormat dateTimeFormat, may be {@code null} 
   *     in which case {@code defaultDateTimeFormat} is used.
   * @return the string which expresses the value of the cell.
   * @throws BizLogicAppException BizLogicAppException
   */
  public @Nullable String getStringFromCell(@Nullable Cell cell, DateTimeFormatter dateTimeFormat)
      throws BizLogicAppException {
    if (dateTimeFormat == null) {
      dateTimeFormat = defaultDateTimeFormat;
    }

    String cellTypeString = null;
    if (cell == null) {
      cellTypeString = "(cell is null)";

    } else {
      cellTypeString = cell.getCellType().toString();
    }

    detailLog.debug("-----");
    detailLog.debug("cellType: " + cellTypeString);

    String value = internalGetStringFromCell(cell, dateTimeFormat);

    detailLog.debug("value: " + (value == null ? "(null)" : value));

    return value;
  }

  /**
   * Returns the value of the cell.
   * 
   * @param cell cell, may be {@code null}.
   * @param dateTimeFormat dateTimeFormat
   * @return the string value of the cell, may be {@code null}.
   * @throws BizLogicAppException BizLogicAppException
   */
  private @Nullable String internalGetStringFromCell(@Nullable Cell cell,
      DateTimeFormatter dateTimeFormat) throws BizLogicAppException {

    // cellがnullの場合もnoDataStringを返す
    if (cell == null) {
      return noDataString;
    }

    CellType cellType = cell.getCellType();

    if (cellType == CellType.FORMULA) {
      return internalGetStringFromCellOtherThanFormulaCellType(cell,
          cell.getCachedFormulaResultType(), dateTimeFormat);

    } else {
      return internalGetStringFromCellOtherThanFormulaCellType(cell, cell.getCellType(),
          dateTimeFormat);
    }
  }

  /**
   * Returns the value of the argument cell in {@code String} format.
   * 
   * <p>Usually the second argument {@code cellType} is equal to {@code cell.getCellType()} 
   *     with the 1st argument {@code cell}.<br>
   *     But when the cellType of the 1st argument {@code cell} is {@code formula}, 
   *     the 2nd argumenet is {@code cell.getCachedFormulaResultType()},
   *     the resulting cellType of the formula cell.</p>
   * 
   * @param cell cell
   * @param cellType cellType
   * @param dateTimeFormat dateTimeFormat
   * @return String value of the cell, may be null when the value in the cell is empty.
   * @throws BizLogicAppException BizLogicAppException
   */
  private @Nullable String internalGetStringFromCellOtherThanFormulaCellType(@Nonnull Cell cell,
      @Nullable CellType cellType, DateTimeFormatter dateTimeFormat) throws BizLogicAppException {

    // poiでは、セルが空欄なら、表示形式に関係なくBLANKというcellTypeになるため、それで判別してから文字を返す
    if (cellType == CellType.BLANK) {
      return noDataString;

    } else if (cellType == CellType.STRING) {
      // 文字列形式
      return getNoDataStringIfNoData(cell.getStringCellValue());

    } else if (cellType == CellType.NUMERIC) {
      // 数値 / 日付形式
      DataFormatter fmter = new DataFormatter();
      Format fmt = fmter.createFormat(cell);

      // fmtにより細かい表示形式の判別が可能
      detailLog.debug("Format: " + ((fmt == null) ? "(null)" : fmt.getClass().getSimpleName()));
      
      if (fmt == null) {
        // DataFormatter#createFormat(Cell) is nullable.
        // In that case return the value without formatting.
        return Double.toString(cell.getNumericCellValue());

      } else if (fmt instanceof ExcelStyleDateFormatter) {
        // format: date and time
        return cell.getLocalDateTimeCellValue().format(dateTimeFormat);

      } else {
        // 表示形式：数値
        String fmtVal = fmt.format(cell.getNumericCellValue());

        String toStrVal = Double.valueOf(cell.getNumericCellValue()).toString();

        // toStrValとfmtValが異なる場合はwarningをあげておく
        boolean warning = false;
        if (!fmtVal.equals(toStrVal)) {

          // fmtValが整数の場合
          if (!fmtVal.contains(".")) {
            // fmtValが整数、toStrValが、fmtVal + ".0"の場合は、問題なし。それ以外の場合はwarningを出す
            if (!(toStrVal.endsWith(".0")
                && fmtVal.equals(toStrVal.substring(0, toStrVal.indexOf("."))))) {
              warning = true;
            }
          }
        }

        if (warning && !suppressesWarnLog) {
          detailLog.warn("The number actual and displayed in excel differs. actual: " + toStrVal
              + "、displayed: " + fmtVal);
        }

        return fmtVal;
      }

    } else if (cellType == CellType.ERROR) {
      // We've got this when the cell says "#NUM!" in excel.
      throw new BizLogicAppException("jp.ecuacion.util.poi.excel.CellContainsError.message",
          cell.getRow().getSheet().getSheetName(), cell.getAddress().formatAsString(),
          cell.getAddress().formatAsR1C1String());

    } else {
      throw new RuntimeException("cell type not found. cellType: " + cellType.toString());
    }
  }

  /**
   * Opens the excel file and returns {@code Workbook} object.
   * 
   * @param filePath filePath
   * @return workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public Workbook openForRead(String filePath) throws EncryptedDocumentException, IOException {
    return WorkbookFactory.create(new File(filePath), null, true);
  }

  /**
   * Gets ready to read table data.
   * 
   * @param ignoresColumnSizeSetInReader It is {@code true} means 
   *     that even if the reader determines the column size,
   *     this method obtains all the columns as long as the header column exists.
   */
  public <T> ContextContainer getReadyToReadTableData(ExcelTableReader<T> reader, Workbook workbook,
      String sheetName, int tableStartColumnNumber,
      Integer numberOfHeaderLinesIfReadsHeaderOnlyOrNull, boolean ignoresColumnSizeSetInReader)
      throws BizLogicAppException {
    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("starting to read excel file.");
    detailLog.debug("sheet name :" + sheetName);

    Sheet sheet = workbook.getSheet(sheetName);

    if (sheet == null) {
      throw new BizLogicAppException("jp.ecuacion.util.poi.excel.SheetNotExist.message", sheetName);
    }

    Integer tableRowSize =
        numberOfHeaderLinesIfReadsHeaderOnlyOrNull == null ? reader.getTableRowSize()
            : numberOfHeaderLinesIfReadsHeaderOnlyOrNull;

    // poiBasis means the top-left position is (0, 0)
    // while tableStartRowNumber / tableStartColumnNumber >= 1.
    final int poiBasisTableStartRowNumber =
        reader.getPoiBasisDeterminedTableStartRowNumber(sheet, tableStartColumnNumber);
    final int poiBasisTableStartColumnNumber = reader.getPoiBasisDeterminedTableStartColumnNumber();
    ContextContainer context =
        new ContextContainer(sheet, poiBasisTableStartRowNumber, poiBasisTableStartColumnNumber,
            tableRowSize, reader.getTableColumnSize(sheet, poiBasisTableStartRowNumber,
                poiBasisTableStartColumnNumber, ignoresColumnSizeSetInReader));

    return context;
  }

  /**
   * Provides common procedure for read one line of a table.
   *
   * @throws BizLogicAppException BizLogicAppException
   */
  public <T> List<T> readTableLine(ExcelTableReader<T> reader, ContextContainer context,
      int rowNumber) throws BizLogicAppException {
    detailLog.debug(LogUtil.PARTITION_MEDIUM);
    detailLog.debug("row number：" + rowNumber);

    // 最大行数を超えたらエラー
    if (rowNumber == ContextContainer.max) {
      throw new RuntimeException("'max':" + ContextContainer.max + " exceeded.");
    }

    // 指定行数読み込み完了時の処理
    if (context.tableRowSize != null
        && rowNumber >= context.poiBasisTableStartRowNumber + context.tableRowSize) {
      throw new LoopBreakException();
    }

    List<T> colList = new ArrayList<>();
    // excel上でtable範囲が終わった場合は、明示的に「row = null」となる。その場合、対象行は空行扱い。
    boolean isEmptyRow = true;

    // excelデータを読み込み。
    for (int j = context.poiBasisTableStartColumnNumber; j < context.poiBasisTableStartColumnNumber
        + context.tableColumnSize; j++) {

      if (reader.isVerticalAndHorizontalOpposite()) {
        Row row = context.sheet.getRow(j);
        if (row == null || row.getCell(rowNumber) == null) {
          colList.add(null);

        } else {
          Cell cell = row.getCell(rowNumber);
          T cellData = reader.getCellData(cell, j + 1);
          colList.add(cellData);
        }

      } else {
        Row row = context.sheet.getRow(rowNumber);
        if (row == null || row.getCell(j) == null) {
          colList.add(null);

        } else {
          Cell cell = row.getCell(j);
          T cellData = reader.getCellData(cell, j + 1);
          colList.add(cellData);
        }
      }
    }

    // 空行チェック。全項目が空欄の場合は空行を意味する。
    for (T colData : colList) {
      if (!reader.isCellDataEmpty(colData)) {
        isEmptyRow = false;
        break;
      }
    }

    // 空行時の処理
    if (isEmptyRow) {
      detailLog.debug("(no data in the line)");
      detailLog.debug(LogUtil.PARTITION_MEDIUM);

      if (context.tableRowSize == null) {
        // 空行発生時に読み込み終了の場合
        throw new LoopBreakException();

      } else {
        // 空行は、それとわかるように要素数ゼロのlistとしておく
        return new ArrayList<>();
      }
    }

    return colList;
  }

  /**
   * Sets {@code suppressesWarnLog}.
   * 
   * @param suppressesWarnLog suppressesWarnLog
   * @return {@code ExcelTableReader<T>}
   */
  public ExcelReadUtil suppressesWarnLog(boolean suppressesWarnLog) {
    this.suppressesWarnLog = suppressesWarnLog;
    return this;
  }
}
