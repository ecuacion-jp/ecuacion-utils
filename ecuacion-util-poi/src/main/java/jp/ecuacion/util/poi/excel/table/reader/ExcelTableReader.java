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
package jp.ecuacion.util.poi.excel.table.reader;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.constraints.Min;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.BeanValidationUtil;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Is a parent of excel table reader classes.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTableReader<T> extends ExcelTable<T>
    implements IfExcelTableReader<T> {

  private DetailLogger detailLog = new DetailLogger(this);
  private ExcelReadUtil readUtil = new ExcelReadUtil();

  /**
   * Is the row size of the table. 
   * 
   * <p>It's equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableRowSize} is 
   *     decided for the program to find an empty row.<br>
   *     When the table has a header, the row size includes the header line,
   */
  @Min(1)
  protected Integer tableRowSize;
  
  /**
   * Is the column size of the table.
   * 
   * <p>It's equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableColumnSize} is 
   *     decided by the length of the header. 
   *     Empty header cell means it's the end of the header.<br>
   *     When the table has a header, the row size includes the header line,   */
  @Min(1)
  protected Integer tableColumnSize;

  /**
   * Constructs a new instance with the sheet name, the position and the size of the excel table.
   * 
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   * @param tableRowSize See {@link ExcelTableReader#tableRowSize}.
   * @param tableColumnSize See {@link ExcelTableReader#tableColumnSize}.
   */
  public ExcelTableReader(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber, @Nullable Integer tableRowSize,
      @Nullable Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.tableRowSize = tableRowSize;
    this.tableColumnSize = tableColumnSize;

    // Validate the input values.
    Set<ConstraintViolation<ExcelTableReader<T>>> violationSet =
        new BeanValidationUtil().validate(this, Locale.getDefault());
    if (violationSet != null && violationSet.size() > 0) {

      throw new RuntimeException("Validation failed at TableReader constructor.");
    }
  }

  /**
   * Reads a table data in an excel file at {@code excelPath} 
   *     and Return it in the form of {@code List<List<String>>}.
   * 
   * <p>The internal {@code List<String>} stores data in one line.<br>
   * The external {@code List} stores lines of {@code List<String>}.</p>
   *
   * @throws IOException IOException
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  @Nonnull
  public List<List<T>> read(@RequireNonnull String excelPath)
      throws EncryptedDocumentException, AppException, IOException {
    List<List<T>> rtnList = readTableValues(excelPath);

    // ヘッダ行のチェック。同時にヘッダ行はexcelTableDataListからremoveしておき、returnするデータには含めない
    List<List<String>> headerList = updateAndGetHeaderList(rtnList);

    validateHeader(headerList);

    return rtnList;
  }

  /*
   * get Table Values in the form of the list of the lists.
   */
  @Nonnull
  private List<List<T>> readTableValues(@RequireNonnull String excelPath)
      throws AppException, EncryptedDocumentException, IOException {
    
    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("starting to read excel file.");
    detailLog.debug("file name  :" + excelPath);
    detailLog.debug("sheet name :" + getSheetName());

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

    // データを取得
    // 2重のlistに格納する
    List<List<T>> rowList = new ArrayList<>();
    int max = 10000;
    for (int rowNumber = poiBasisTableStartRowNumber; rowNumber <= max; rowNumber++) {
      detailLog.debug(LogUtil.PARTITION_MEDIUM);
      detailLog.debug("row number：" + rowNumber);

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
          T cellData = getCellData(cell, j + 1);
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

    detailLog.debug("finishing to read excel file. sheet name :" + getSheetName());
    detailLog.debug(LogUtil.PARTITION_LARGE);

    return rowList;
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

      if (isCellDataEmpty(getCellData(cell, tableStartColumnNumber + columnNumber + 1))) {
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
    detailLog.debug("(no data in the line)");
    detailLog.debug(LogUtil.PARTITION_MEDIUM);
  }

  /**
   * sets {@code tableColumnSize}.
   * 
   * <p>tableColumnSize can be set by the header length, 
   *     but the instance method cannot be called from constructors so the setter is needed.</p>
   *     
   * <p>This method set the final value of the column size, 
   *     so the argument is not {@code Integer}, but primitive {@code int}.
   * 
   * @param tableColumnSize tableColumnSize.
   */
  public void setTableColumnSize(int tableColumnSize) {
    this.tableColumnSize = tableColumnSize;
  }

  @Override
  public ExcelReadUtil getExcelReadUtil() {
    return readUtil;
  }
}
