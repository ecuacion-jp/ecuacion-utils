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
import java.util.Objects;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.reader.IfExcelTableReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Stores properties in an excel table.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTable<T> implements IfExcelTable<T> {

  /**
   * Is the sheet name of the excel file.
   */
  @NotNull
  @Nonnull
  protected String sheetName;

  /**
   * Is the row number from which the table starts.
   * 
   * <p>The minimum value is {@code 1}, 
   *     which means the table starts at the first line of the excel sheet.</p>
   *     
   * <p>{@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableStartRowNumber} is 
   *     decided by the far left header value of the table.</p>
   *     
   * <p>The header value is obtained from 
   *     {@link IfExcelTable#getFarLeftAndTopHeaderLabel()}.</p>
   */
  @Nullable
  @Min(1)
  protected Integer tableStartRowNumber;

  /**
   * Is the column number from which the table starts.
   * 
   * <p>The minimum value is {@code 1}, 
   *     which means the table starts at the far left column of the excel sheet.</p>
   * 
   * <p>{@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is not acceptable.<br>
   *     (Its data type is primitive {@code int}, so it can't have {@code null} anyway.)
   */
  @Min(1)
  protected int tableStartColumnNumber;

  protected boolean ignoresAdditionalColumnsOfHeaderData;

  /**
   * Constructs a new instance with the sheet name, the position and the size of the excel table.
   * 
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public ExcelTable(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    this.sheetName = ObjectsUtil.paramRequireNonNull(sheetName);
    this.tableStartRowNumber = tableStartRowNumber;
    this.tableStartColumnNumber = tableStartColumnNumber;
  }

  public @Nonnull String getSheetName() {
    return ObjectsUtil.returnRequireNonNull(sheetName);
  }

  /**
   * Returns the row number at which the table starts.
   * 
   * <p>The minimum value of {@code tableStartRowNumber} is zero
   *     bacause the top-left of the excel sheet is (1, 1) in R1C1 format, 
   *     but since apache poi specifies the the top-left of the excel sheet is (0, 0),
   *     this method returns the poi-based row number.</p>
   * 
   * <p>When {@code tableStartRowNumber} is set to {@code null}, 
   *     this method will find the string designated with 
   *     {@link IfExcelTableReader#getFarLeftAndTopHeaderLabel()} from the top row 
   *     in the column number of {@code excelBasisTableStartColumnNumber}.</p>
   * 
   * @param sheet excel sheet
   * @return the row number the table starts, greater than or equal to {@code 1}.
   */
  public int getPoiBasisDeterminedTableStartRowNumber(@RequireNonnull Sheet sheet,
      int excelBasisTableStartColumnNumber) {
    ObjectsUtil.paramRequireNonNull(sheet);

    if (tableStartRowNumber != null) {
      return Objects.requireNonNull(tableStartRowNumber) - 1;
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
      Cell cell = row.getCell(excelBasisTableStartColumnNumber - 1);
      // cellがnullになる場合もあるのでその場合はskip
      if (cell == null) {
        continue;
      }

      // 文字列の取得
      String value = cell.getStringCellValue();

      if (value.equals(getFarLeftAndTopHeaderLabel())) {
        // iはプログラム上の行（ゼロ始まり）だが、getTableStartRowNumber()としては左上が(1, 1)として返すルールなので1をプラスして返す
        return i;
      }
    }

    // ここまでくるということは、signStringがなかったということ。異常終了
    throw new RuntimeException("シート「" + sheet.getSheetName() + "」に文字列「"
        + getFarLeftAndTopHeaderLabel() + "」が" + tableStartColumnNumber + "番目の列に存在しません。終了します。");
  }

  /**
   * Returns tableStartColumnNumber.
   * 
   * <p>The minimum value of {@code tableStartColumnNumber} is zero
   *     bacause the top-left of the excel sheet is (1, 1) in R1C1 format, 
   *     but since apache poi specifies the the top-left of the excel sheet is (0, 0),
   *     this method returns the poi-based row number.</p>
   *     
   * @return the column number the table starts
   */
  public int getPoiBasisDeterminedTableStartColumnNumber() {
    return tableStartColumnNumber - 1;
  }

  /**
   * Stores context data.
   */
  public static class ContextContainer {
    public final Sheet sheet;
    public final int poiBasisTableStartRowNumber;
    public final int poiBasisTableStartColumnNumber;
    public final Integer tableRowSize;
    public final Integer tableColumnSize;

    public static final int max = 10000;

    /**
     * Constructs a new instance.
     * 
     * @param sheet sheet
     * @param poiBasisTableStartColumnNumber poiBasisTableStartColumnNumber
     * @param poiBasisTableStartRowNumber poiBasisTableStartRowNumber
     * @param tableColumnSize tableColumnSize
     * @param tableRowSize tableRowSize
     */
    public ContextContainer(Sheet sheet, int poiBasisTableStartRowNumber,
        int poiBasisTableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {
      this.sheet = sheet;
      this.poiBasisTableStartRowNumber = poiBasisTableStartRowNumber;
      this.poiBasisTableStartColumnNumber = poiBasisTableStartColumnNumber;
      this.tableRowSize = tableRowSize;
      this.tableColumnSize = tableColumnSize;
    }
  }

  @Override
  public boolean ignoresAdditionalColumnsOfHeaderData() {
    return ignoresAdditionalColumnsOfHeaderData;
  }
}
