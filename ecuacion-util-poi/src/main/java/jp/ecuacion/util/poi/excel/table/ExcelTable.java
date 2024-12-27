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
package jp.ecuacion.util.poi.excel.table;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotNull;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.reader.core.IfExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Stores properties an excel table has.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTable<T> implements IfExcelTable<T> {

  @NotNull
  @Nonnull
  protected String sheetName;
  @Min(1)
  protected Integer tableStartRowNumber;
  @Min(1)
  protected int tableStartColumnNumber;
  
  /**
   * Constructs a new instance with the sheet name, the position and the size of the excel table.
   * 
   * @param sheetName the sheet name of the excel file
   * @param tableStartRowNumber the row number from which the table starts. 
   *     It starts with {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableStartRowNumber} is 
   *     decided by the far left header value of the table.
   *     The header value is obtained from 
   *     {@link IfExcelTable#getFarLeftHeaderLabel()}.
   * @param tableStartColumnNumber the column number from which the table starts.
   *     It starts with {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is not acceptable. 
   *     (Its data type is primitive {@code int}, so it can't have {@code null} anyway.)
   */
  public ExcelTable(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    this.sheetName = ObjectsUtil.paramRequireNonNull(sheetName);
    this.tableStartRowNumber = tableStartRowNumber;
    this.tableStartColumnNumber = tableStartColumnNumber;
  }

  /**
   * See {@link IfExcelTable#getSheetName()}.
   */
  public @Nonnull String getSheetName() {
    return ObjectsUtil.returnRequireNonNull(sheetName);
  }

  /**
   * Returns the row number the table starts.
   * 
   * <p>When {@code tableStartRowNumber} is set to {@code null}, 
   *     this method will find the string designated with 
   *     {@link IfExcelTableReader#getFarLeftHeaderLabel()} from the top row 
   *     in the column number of {@code tableStartColumnNumber}.</p>
   * 
   * @param sheet excel sheet
   * @return the row number the table starts, greater than or equal to {@code 1}.
   */
  protected int getPoiBasisDeterminedTableStartRowNumber(@RequireNonnull Sheet sheet) {
    ObjectsUtil.paramRequireNonNull(sheet);

    if (tableStartRowNumber != null) {
      return tableStartRowNumber - 1;
    }

    // 以下、tableStartRowNumberを動的に決める必要がある場合の処理

    // A列に特定の文字列があることを確認
    for (int i = 0; i < 100; i++) {
      // i行目
      Row row = sheet.getRow(i);
      // 空行がnullになる場合もあるのでその場合はskip
      if (row == null) {
        continue;
      }

      // 0番目のセル
      Cell cell = row.getCell(0);
      // cellがnullになる場合もあるのでその場合はskip
      if (cell == null) {
        continue;
      }

      // 文字列の取得
      String value = cell.getStringCellValue();

      if (value.equals(getFarLeftHeaderLabel())) {
        // iはプログラム上の行（ゼロ始まり）だが、getTableStartRowNumber()としては左上が(1, 1)として返すルールなので1をプラスして返す
        return i;
      }
    }

    // ここまでくるということは、signStringがなかったということ。異常終了
    throw new RuntimeException("シート「" + sheet.getSheetName() + "」に文字列「" + getFarLeftHeaderLabel()
        + "」が" + tableStartColumnNumber + "番目の列に存在しません。終了します。");
  }

  /**
   * Returns tableStartColumnNumber.
   * 
   * @return the column number the table starts
   */
  protected int getPoiBasisDeterminedTableStartColumnNumber() {
    return tableStartColumnNumber - 1;
  }
}
