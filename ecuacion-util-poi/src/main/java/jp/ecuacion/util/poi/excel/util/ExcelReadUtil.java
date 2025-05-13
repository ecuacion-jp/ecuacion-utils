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
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;
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
   * @throws ExcelAppException ExcelAppException
  */
  public @Nullable String getStringFromCell(@Nullable Cell cell) throws ExcelAppException {
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
   * @throws ExcelAppException ExcelAppException
   */
  public @Nullable String getStringFromCell(@Nullable Cell cell, DateTimeFormatter dateTimeFormat)
      throws ExcelAppException {
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
   * @throws ExcelAppException ExcelAppException
   */
  private @Nullable String internalGetStringFromCell(@Nullable Cell cell,
      DateTimeFormatter dateTimeFormat) throws ExcelAppException {

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
   * @throws ExcelAppException ExcelAppException
   */
  private @Nullable String internalGetStringFromCellOtherThanFormulaCellType(@Nonnull Cell cell,
      @Nullable CellType cellType, DateTimeFormatter dateTimeFormat) throws ExcelAppException {

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
      throw new ExcelAppException("jp.ecuacion.util.poi.excel.CellContainsError.message",
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
  public static Workbook openForRead(String filePath)
      throws EncryptedDocumentException, IOException {
    return WorkbookFactory.create(new File(filePath), null, true);
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
