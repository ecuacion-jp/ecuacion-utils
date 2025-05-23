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

import jakarta.annotation.Nullable;
import java.io.File;
import java.io.IOException;
import java.text.Format;
import java.time.format.DateTimeFormatter;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.PropertyFileUtil.Arg;
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
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

  private static DetailLogger detailLog = new DetailLogger(ExcelReadUtil.class);

  /* 空文字オブジェクトが複数作られると非効率なので定義しておく。 */
  private static final String EMPTY_STRING = "";

  private static DateTimeFormatter defaultDateTimeFormat =
      DateTimeFormatter.ofPattern("yyyy-MM-dd");

  /**
   * Prevents other classes from instantiating it.
   */
  private ExcelReadUtil() {}

  /** 
   * Returns proper {@code NoDataString} value if the argument value is 
   *     {@code null} or {@code ""}, otherwize returns the argument value.
   * 
   * @param value value, may be {@code null}.
   * @return the argument value or the empty string designated by {@code noDataString} 
   *     when the argument value is empty.
   */
  public static @Nullable String getNoDataStringIfNoData(@Nullable String value,
      String noDataString) {
    if (value == null || value.equals(EMPTY_STRING)) {
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
  public static @Nullable String getStringFromCell(@Nullable Cell cell) throws ExcelAppException {
    return getStringFromCell(cell, null, defaultDateTimeFormat);
  }

  /**
   * Returns {@code String} format cell value
   * in spite of the format or value kind of the cell.
   *
   * @param filename Used for error message, 
   *     may be {@code null} in which case the error message shows no filename.
   * @param cell the cell of the excel file
   * @return the string which expresses the value of the cell.
   * @throws ExcelAppException ExcelAppException
   */
  public static @Nullable String getStringFromCell(@Nullable Cell cell, @Nullable String filename)
      throws ExcelAppException {
    return getStringFromCell(cell, filename, defaultDateTimeFormat);
  }

  /**
   * Returns {@code String} format cell value 
   *     in spite of the format or value kind of the cell.
   * 
   * <p>return value when row is null or cell is null, etc... is null.</p>
   * 
   * @param filename Used for error message, 
   *     may be {@code null} in which case the error message shows no filename.
   * @param cell the cell of the excel file
   * @param dateTimeFormat dateTimeFormat, may be {@code null} 
   *     in which case {@code defaultDateTimeFormat} is used.
   * @return the string which expresses the value of the cell.
   * @throws ExcelAppException ExcelAppException
   */
  public static @Nullable String getStringFromCell(@Nullable Cell cell, @Nullable String filename,
      DateTimeFormatter dateTimeFormat) throws ExcelAppException {
    return getStringFromCell(cell, filename, dateTimeFormat, null);
  }

  /**
   * Returns {@code String} format cell value 
   *     in spite of the format or value kind of the cell.
   * 
   * @param filename Used for error message, 
   *     may be {@code null} in which case the error message shows no filename.
   * @param cell the cell of the excel file
   * @param dateTimeFormat dateTimeFormat, may be {@code null} 
   *     in which case {@code defaultDateTimeFormat} is used.
   * @param noDataString return value when row is null or cell is null, etc...
   * @return the string which expresses the value of the cell.
   * @throws ExcelAppException ExcelAppException
   */
  public static @Nullable String getStringFromCell(@Nullable Cell cell, @Nullable String filename,
      DateTimeFormatter dateTimeFormat, String noDataString) throws ExcelAppException {
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

    String value = internalGetStringFromCell(cell, filename, dateTimeFormat, noDataString);

    detailLog.debug("value: " + (value == null ? "(null)" : value));

    return value;
  }

  /**
   * Returns the value of the cell.
   * 
   * @param filename Used for error message, 
   *     may be {@code null} in which case the error message shows no filename.
   * @param cell cell, may be {@code null}.
   * @param dateTimeFormat dateTimeFormat
   * @return the string value of the cell, may be {@code null}.
   * @throws ExcelAppException ExcelAppException
   */
  private static @Nullable String internalGetStringFromCell(@Nullable Cell cell,
      @Nullable String filename, DateTimeFormatter dateTimeFormat, String noDataString)
      throws ExcelAppException {

    // cellがnullの場合もnoDataStringを返す
    if (cell == null) {
      return noDataString;
    }

    CellType cellType = cell.getCellType();

    if (cellType == CellType.FORMULA) {
      return internalGetStringFromCellOtherThanFormulaCellType(cell, filename,
          cell.getCachedFormulaResultType(), noDataString, dateTimeFormat);

    } else {
      return internalGetStringFromCellOtherThanFormulaCellType(cell, filename, cell.getCellType(),
          noDataString, dateTimeFormat);
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
   * @param filename Used for error message, 
   *     may be {@code null} in which case the error message shows no filename.
   * @param cell cell
   * @param cellType cellType
   * @param dateTimeFormat dateTimeFormat
   * @return String value of the cell, may be null when the value in the cell is empty.
   * @throws ExcelAppException ExcelAppException
   */
  private static @Nullable String internalGetStringFromCellOtherThanFormulaCellType(
      @RequireNonnull Cell cell, @Nullable String filename, @Nullable CellType cellType,
      String noDataString, DateTimeFormatter dateTimeFormat) throws ExcelAppException {

    // poiでは、セルが空欄なら、表示形式に関係なくBLANKというcellTypeになるため、それで判別してから文字を返す
    if (cellType == CellType.BLANK) {
      return noDataString;

    } else if (cellType == CellType.STRING) {
      // 文字列形式
      return getNoDataStringIfNoData(cell.getStringCellValue(), noDataString);

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

        if (warning) {
          detailLog.debug("The number actual and displayed in excel differs. actual: " + toStrVal
              + "、displayed: " + fmtVal);
        }

        return fmtVal;
      }

    } else if (cellType == CellType.ERROR) {
      // We've got this when the cell says "#NUM!" in excel.
      throw new ExcelAppException("jp.ecuacion.util.poi.excel.CellContainsError.message",
          ArrayUtils.addAll(
              Arg.strings(cell.getRow().getSheet().getSheetName(),
                  cell.getAddress().formatAsString()),
              StringUtils.isEmpty(filename) ? Arg.strings("", "")
                  : new Arg[] {Arg.message("jp.ecuacion.util.poi.common.messageItemSeparator"),
                      Arg.message("jp.ecuacion.util.poi.common.filename", filename)}));

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
}
