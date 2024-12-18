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
package jp.ecuacion.util.poi.read.string.reader.internal;

import jakarta.validation.constraints.Min;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.exception.checked.MultipleAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.BeanValidationUtil;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.util.PoiReadUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public abstract class PoiStringTableReader {
  protected PoiReadUtil readUtil;
  protected DetailLogger detailLog = new DetailLogger(this);

  public PoiStringTableReader() {
    readUtil = new PoiReadUtil();
  }

  public PoiStringTableReader(NoDataString noDataString) {
    readUtil = new PoiReadUtil(noDataString);
  }

  protected abstract String getSheetName();

  /** tableのヘッダが存在する行の行番号を指す。excel行1行目にヘッダがある場合は"1"。"0"は設定不可。 */
  protected abstract int getTableStartRowNumber(Sheet sheet);

  /**
   * 読み込む行数。ヘッダ行を含めた行数を設定。"0"は設定不可。 nullを返すことも可能で、その場合は、全項目が空欄の行が発生した時点で読み込み終了。
   */
  protected abstract Integer getTableRowSize();

  /** tableがはじまる列番号を指す。1列目からtableがはじまる場合は"1"。"0"は設定不可。 */
  protected abstract int getTableStartColumnNumber();

  /** tableの列数。1列のみ存在するtable（現実あまりないとは思うが）の場合は1。"0"は設定不可。 */
  protected abstract int getTableStartColumnSize();

  /**
   * excelを読み込みList&lt;List&lt;String&gt;&gt;（内側のList&lt;String&gt;に、1行内の複数列分の情報を格納）。 子クラスにて実装。
   */
  protected abstract List<List<String>> getTableValues(String excelPath)
      throws AppException, EncryptedDocumentException, IOException;

  protected List<List<String>> getTableValuesCommon(String excelPath)
      throws AppException, EncryptedDocumentException, IOException {

    Workbook excel = WorkbookFactory.create(new File(excelPath), null, true);
    Sheet sheet = excel.getSheet(getSheetName());

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", excelPath, getSheetName());
    }

    // ...StartRowNumberは、get〜がエクセルの左上を1行目・1列目としているのに対し、プログラム上では(0, 0)とする必要があるのでずらす
    final int tableStartRowNumber = getTableStartRowNumber(sheet) - 1;
    // nullの場合は空行があった時点で読み込み終了
    final Integer tableRowSize = getTableRowSize();
    final int tableStartColumnNumber = getTableStartColumnNumber() - 1;
    final int tableStartColumnSize = getTableStartColumnSize();

    checkNumbers(tableStartRowNumber, tableRowSize, tableStartColumnNumber, tableStartColumnSize);

    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("excelファイル読み取り処理開始");
    detailLog.debug("ファイル名：" + excelPath);
    detailLog.debug("sheet名：" + getSheetName());

    // データを取得
    // 2重のlistに格納する
    List<List<String>> rowList = new ArrayList<>();
    int max = 10000;
    for (int rowNumber = tableStartRowNumber; rowNumber <= max; rowNumber++) {
      detailLog.debug(LogUtil.PARTITION_MEDIUM);
      detailLog.debug("処理対象行：" + rowNumber);

      // 最大行数を超えたらエラー
      if (rowNumber == max) {
        throw new RuntimeException("'max':" + max + " exceeded.");
      }

      // 指定行数読み込み完了時の処理
      if (tableRowSize != null && rowNumber >= tableStartRowNumber + tableRowSize) {
        break;
      }

      Row row = sheet.getRow(rowNumber);
      List<String> colList = new ArrayList<>();

      // excel上でtable範囲が終わった場合は、明示的に「row = null」となる。その場合、対象行は空行扱い。
      boolean isEmptyRow = true;
      if (row != null) {
        // excelデータを読み込み。
        for (int j = tableStartColumnNumber; j < tableStartColumnNumber
            + tableStartColumnSize; j++) {
          Cell cell = row.getCell(j);
          String cellString = readUtil.getStringFromCell(cell);
          colList.add(cellString);
        }

        // 空行チェック。全項目が空欄の場合は空行を意味する。
        for (String colString : colList) {
          if (!StringUtils.isEmpty(colString)) {
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

  private void logNoMoreLines() {
    detailLog.debug("（対象行にデータなし）");
    detailLog.debug(LogUtil.PARTITION_MEDIUM);
  }

  protected void checkNumbers(int tableStartRowNumber, Integer tableRowSize,
      int tableStartColumnNumber, int tableStartColumnSize) throws MultipleAppException {
    Numbers obj = new Numbers(tableStartRowNumber, tableRowSize, tableStartColumnNumber,
        tableStartColumnSize);
    new BeanValidationUtil().validateThenThrow(obj);
  }

  /** 
   * validation checkのみの目的で使用するclass. fieldはspotbug対策のためだけにpublic化。
   * 
   */
  private static class Numbers {
    @Min(0)
    public int tableStartRowNumber;
    @Min(1)
    public Integer tableRowSize;
    @Min(0)
    public int tableStartColumnNumber;
    @Min(1)
    public int tableStartColumnSize;

    public Numbers(int tableStartRowNumber, Integer tableRowSize, int tableStartColumnNumber,
        int tableStartColumnSize) {
      this.tableStartRowNumber = tableStartRowNumber;
      this.tableRowSize = tableRowSize;
      this.tableStartColumnNumber = tableStartColumnNumber;
      this.tableStartColumnSize = tableStartColumnSize;
    }
  }
}
