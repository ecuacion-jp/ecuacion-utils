/*
 * Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License
 * is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied. See the License for the specific language governing permissions and limitations under
 * the License.
 */
package jp.ecuacion.util.poi.excel.util;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.text.Format;
import java.text.SimpleDateFormat;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.ExcelStyleDateFormatter;

/**
 * Provides excel-related {@code apache POI} utility methods.
 */
public class ExcelReadUtil {

  private DetailLogger detailLog = new DetailLogger(this);

  /* 空文字オブジェクトが複数作られると非効率なので定義しておく。 */
  private static final String EMPTY_STRING = "";

  private String noDataString;

  private String dateFormat = "yyyy-MM-dd";

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
   * Sets dateFormat.
   * 
   * @param dateFormat dateFormat
   * @return ReturnUrlBean (for method chain)
   */
  public ExcelReadUtil dateFormat(String dateFormat) {
    this.dateFormat = dateFormat;
    return this;
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
   *     in spite of the format or value kind of the cell.
   * 
   * @param cell the cell of the excel file
   * @return the string which expresses the value of the cell.
   */
  public @Nullable String getStringFromCell(@Nullable Cell cell) {
    String cellTypeString = null;
    if (cell == null) {
      cellTypeString = "(cell is null)";

    } else {
      cellTypeString = cell.getCellType().toString();
    }

    detailLog.debug("-----");
    detailLog.debug("cellType: " + cellTypeString);

    String value = internalGetStringFromCell(cell);

    detailLog.debug("value: " + (value == null ? "(null)" : value));

    return value;
  }

  /**
   * Returns the value of the cell.
   * 
   * @param cell cell, may be {@code null}.
   * @return the string value of the cell, may be {@code null}.
   */
  private @Nullable String internalGetStringFromCell(@Nullable Cell cell) {

    // cellがnullの場合もnoDataStringを返す
    if (cell == null) {
      return noDataString;
    }

    CellType cellType = cell.getCellType();

    if (cellType == CellType.FORMULA) {
      return internalGetStringFromCellOtherThanFormulaCellType(cell,
          cell.getCachedFormulaResultType());

    } else {
      return internalGetStringFromCellOtherThanFormulaCellType(cell, cell.getCellType());
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
   * @return String value of the cell, may be null when the value in the cell is empty.
   */
  private @Nullable String internalGetStringFromCellOtherThanFormulaCellType(@Nonnull Cell cell,
      @Nullable CellType cellType) {

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
      detailLog.debug("Format: " + fmt.getClass().getSimpleName());
      if (fmt instanceof ExcelStyleDateFormatter) {
        // 表示形式：日付
        // CellStyle style = cell.getCellStyle();

        // 日付の場合のformatは、poi内ではindex番号で管理されており、style.getDataFormat()で取得可能。
        // それに対する実際のformat文字列（yyyy/M/dなど）はgetDataFormatString()で取得。
        // index == 14は、excel（日本語版？）上では「yyyy/M/d」だが、poi内では「m/d/yy」になっている。。。のでOS言語で判断・・
        // String dateFormatString = null;
        // detailLog.debug("dataFormat (index) : " + style.getDataFormat());
        // if (style.getDataFormat() == 14) {
        // Locale locale = Locale.getDefault();
        // dateFormatString =
        // locale.getLanguage().equals("ja") ? "yyyy/M/d" : style.getDataFormatString();
        //
        // } else {
        // detailLog.warn("dataFormat (index) : " + style.getDataFormat());
        // detailLog
        // .warn("The dataFormat other than 14 is not recommended. It may not be correct.");
        // dateFormatString = style.getDataFormatString();
        // }

        // detailLog.debug("dataFormatString(poi original) : " + style.getDataFormatString());
        // detailLog.debug("dataFormatString(corrected) : " + dateFormatString);

        SimpleDateFormat format = new SimpleDateFormat(dateFormat);
        return format.format(cell.getDateCellValue());

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
          detailLog.warn("The number actual and displayed in excel differs. actual: " + toStrVal
              + "、displayed: " + fmtVal);
        }

        return fmtVal;
      }

    } else {
      throw new RuntimeException("cell type not found. cellType: " + cellType.toString());
    }
  }
}
