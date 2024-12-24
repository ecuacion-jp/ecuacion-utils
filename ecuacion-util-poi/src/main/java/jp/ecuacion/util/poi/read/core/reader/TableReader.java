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
package jp.ecuacion.util.poi.read.core.reader;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jakarta.validation.constraints.Min;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.exception.checked.MultipleAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.BeanValidationUtil;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Is a parent of reader classes, which reads excel tables.
 * 
 * @param <T> See {@link IfAbstractFormatTableReader}.
 */
public abstract class TableReader<T> implements IfAbstractFormatTableReader<T> {
  private DetailLogger detailLog = new DetailLogger(this);

  @Nonnull
  private String sheetName;

  private Integer tableStartRowNumber;
  private int tableStartColumnNumber;
  private Integer tableRowSize;
  private Integer tableColumnSize;

  /**
   * Constructs a new instance with the position and the size of the excel table.
   * 
   * @param sheetName the sheet name of the excel file
   * @param tableStartRowNumber the row number from which the table starts, <br>
   *     starts with {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableStartRowNumber} is 
   *     decided by the header value at the top-left of the table.
   * @param tableStartColumnNumber the column number from which the table starts, <br>
   *     starts with {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is not acceptable.
   * @param tableRowSize the row size of the table including the header line, <br>
   *     equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableRowSize} is 
   *     decided when the program finds the empty row.
   * @param tableColumnSize the column size of the table including the header line, <br>
   *     equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableColumnSize} is 
   *     decided by the length of the header. Empty header cell means it's the end of the header.
   */
  public TableReader(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber, @Nullable Integer tableRowSize,
      @Nullable Integer tableColumnSize) {
    this.sheetName = ObjectsUtil.paramRequireNonNull(sheetName);
    this.tableStartRowNumber = tableStartRowNumber;
    this.tableStartColumnNumber = tableStartColumnNumber;
    this.tableRowSize = tableRowSize;
    this.tableColumnSize = tableColumnSize;
  }
  
  /**
   * Returns the excel sheet name the reader reads.
   * 
   * @return the sheet name
   */
  public @Nonnull String getSheetName() {
    return ObjectsUtil.returnRequireNonNull(sheetName);
  }

  /**
   * Returns the obtained value from the cell.
   * 
   * <p>If you want to get {@code String} value from the cell, it returns the string value.</p>
   * 
   * @param cell cell, may be null.
   * @return the obtained value from the cell
   */
  protected abstract @Nullable T getCellData(@RequireNonnull Cell cell);

  /**
   * Returns whether the value of the cell is empty.
   * 
   * @param cellData cellData
   * @return whether the valule of the cell is empty.
   */
  protected abstract boolean isCellDataEmpty(@Nullable T cellData);

  /**
   * Gets table data list in the form of {@code List<List<String>>}.
   * 
   * <p>header line is not included.</p>
   * excelを読み込みList&lt;List&lt;String&gt;&gt;（内側のList&lt;String&gt;に、1行内の複数列分の情報を格納）。 子クラスにて実装。
   *  
   * @see jp.ecuacion.util.poi.read.string.reader.internal.StringTableReader
   * @throws IOException IOException
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  public List<List<T>> getAndValidateTableValues(String excelPath)
      throws EncryptedDocumentException, AppException, IOException {
    List<List<T>> rtnList = getTableValues(excelPath);

    validateAndUpdate(rtnList);

    return rtnList;
  }

  /*
   * get Table Values in the form of the list of the lists.
   */
  private List<List<T>> getTableValues(String excelPath)
      throws AppException, EncryptedDocumentException, IOException {

    Workbook excel = WorkbookFactory.create(new File(excelPath), null, true);
    Sheet sheet = excel.getSheet(getSheetName());

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", excelPath, getSheetName());
    }

    // poiBasis means the top-left position is (0, 0)
    // while tableStartRowNumber / tableStartColumnNumber >= 1.
    final int poiBasisTableStartColumnNumber = getPoiBasisDeterminedTableStartColumnNumber();
    final int poiBasisTableStartRowNumber = getPoiBasisDeterminedTableStartRowNumber(sheet);
    final int tableColumnSize =
        getTableColumnSize(sheet, poiBasisTableStartRowNumber, poiBasisTableStartColumnNumber);
    // tableRowSizeだけは、この時点では値が確定していない。nullの場合はこの後の処理で空行があった時点で読み込み終了
    final Integer tableRowSize = getTableRowSize();

    checkNumbers(poiBasisTableStartRowNumber, poiBasisTableStartColumnNumber, tableRowSize,
        tableColumnSize);

    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("excelファイル読み取り処理開始");
    detailLog.debug("ファイル名：" + excelPath);
    detailLog.debug("sheet名：" + getSheetName());

    // データを取得
    // 2重のlistに格納する
    List<List<T>> rowList = new ArrayList<>();
    int max = 10000;
    for (int rowNumber = poiBasisTableStartRowNumber; rowNumber <= max; rowNumber++) {
      detailLog.debug(LogUtil.PARTITION_MEDIUM);
      detailLog.debug("処理対象行：" + rowNumber);

      // 最大行数を超えたらエラー
      if (rowNumber == max) {
        throw new RuntimeException("'max':" + max + " exceeded.");
      }

      // 指定行数読み込み完了時の処理
      if (tableRowSize != null && rowNumber >= poiBasisTableStartRowNumber + tableRowSize) {
        break;
      }

      Row row = sheet.getRow(rowNumber);
      List<T> colList = new ArrayList<>();

      // excel上でtable範囲が終わった場合は、明示的に「row = null」となる。その場合、対象行は空行扱い。
      boolean isEmptyRow = true;
      if (row != null) {
        // excelデータを読み込み。
        for (int j = poiBasisTableStartColumnNumber; j < poiBasisTableStartColumnNumber
            + tableColumnSize; j++) {
          Cell cell = row.getCell(j);
          T cellData = getCellData(cell);
          colList.add(cellData);
        }

        // 空行チェック。全項目が空欄の場合は空行を意味する。
        for (T colData : colList) {
          if (!isCellDataEmpty(colData)) {
            isEmptyRow = false;
            break;
          }
        }
      }

      // 空行時の処理
      if (isEmptyRow) {
        logNoMoreLines();
        if (tableRowSize == null) {
          // 空行発生時に読み込み終了の場合
          break;

        } else {
          // 空行は、それとわかるように要素数ゼロのlistとしておく
          rowList.add(new ArrayList<>());
          continue;
        }
      }

      rowList.add(colList);
    }

    excel.close();

    detailLog.debug("（excelファイル読み取り処理正常終了）sheet名：" + getSheetName());
    detailLog.debug(LogUtil.PARTITION_LARGE);

    return rowList;
  }

  /**
   * Returns the row number the table starts.
   * 
   * <p>When {@code tableStartRowNumber} is set to {@code null}, 
   *     this method will find the string designated with 
   *     {@code getHeaderLabelToDecideTableStartRowNumber()} from the top row 
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

      if (value.equals(getHeaderLabelToDecideTableStartRowNumber())) {
        // iはプログラム上の行（ゼロ始まり）だが、getTableStartRowNumber()としては左上が(1, 1)として返すルールなので1をプラスして返す
        return i;
      }
    }

    // ここまでくるということは、signStringがなかったということ。異常終了
    throw new RuntimeException(
        "シート「" + sheet.getSheetName() + "」に文字列「" + getHeaderLabelToDecideTableStartRowNumber()
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

  /**
   * Returns tableRowSize, may be {@code null}. 
   */
  protected @Nullable Integer getTableRowSize() {
    return tableRowSize;
  }

  /**
   * Returns tableColumnSize, may be {@code null}. 
   * 
   * @param sheet sheet
   * @param poiBasisDeterminedTableStartRowNumber poiBasisDeterminedTableStartRowNumber
   * @param poiBasisDeterminedTableStartColumnNumber poiBasisDeterminedTableStartRowNumber
   * @throws BizLogicAppException BizLogicAppException
   */
  protected @Nonnull Integer getTableColumnSize(@RequireNonnull Sheet sheet,
      int poiBasisDeterminedTableStartRowNumber, int poiBasisDeterminedTableStartColumnNumber)
      throws BizLogicAppException {
    ObjectsUtil.paramRequireNonNull(sheet);

    if (tableColumnSize != null) {
      return Objects.requireNonNull(tableColumnSize);
    }

    // the folloing is the case that tableColumnSize value needs to be analyzed dynamically.

    Row row = sheet.getRow(poiBasisDeterminedTableStartRowNumber);

    // This row is the line with header, which cannot be {@code null}.
    ObjectsUtil.requireNonNull(row);

    int columnNumber = poiBasisDeterminedTableStartColumnNumber;
    while (true) {
      Cell cell = row.getCell(columnNumber);
      // If the cell is null, that means header is end.
      if (cell == null) {
        break;
      }

      if (isCellDataEmpty(getCellData(cell))) {
        break;
      }

      columnNumber++;
    }

    int size = columnNumber - poiBasisDeterminedTableStartColumnNumber;

    if (size == 0) {
      throw new BizLogicAppException("");
    }

    return size;
  }

  private void logNoMoreLines() {
    detailLog.debug("（対象行にデータなし）");
    detailLog.debug(LogUtil.PARTITION_MEDIUM);
  }

  /**
   * sets {@code tableColumnSize}.
   * 
   * <p>tableColumnSize can be set by the header length, 
   *     but the instance method cannot be called from constructors so the setter is needed.</p>
   * 
   * @param tableColumnSize tableColumnSize
   */
  public void setTableColumnSize(Integer tableColumnSize) {
    this.tableColumnSize = tableColumnSize;
  }

  /*
   * Validates 4 numbers.
   */
  private void checkNumbers(Integer tableStartRowNumber, int tableStartColumnNumber,
      Integer tableRowSize, int tableColumnSize) throws MultipleAppException {
    TablePositionAndSize obj = new TablePositionAndSize(tableStartRowNumber, tableStartColumnNumber,
        tableRowSize, tableColumnSize);
    new BeanValidationUtil().validateThenThrow(obj);
  }

  /** 
   * validation checkのみの目的で使用するclass. fieldはspotbug対策のためだけにpublic化。
   */
  private static class TablePositionAndSize {
    @Min(0)
    public Integer tableStartRowNumber;
    @Min(0)
    public int tableStartColumnNumber;
    @Min(1)
    public Integer tableRowSize;
    @Min(1)
    public Integer tableStartColumnSize;

    public TablePositionAndSize(Integer tableStartRowNumber, int tableStartColumnNumber,
        Integer tableRowSize, Integer tableStartColumnSize) {
      this.tableStartRowNumber = tableStartRowNumber;
      this.tableStartColumnNumber = tableStartColumnNumber;
      this.tableRowSize = tableRowSize;
      this.tableStartColumnSize = tableStartColumnSize;
    }
  }
}
