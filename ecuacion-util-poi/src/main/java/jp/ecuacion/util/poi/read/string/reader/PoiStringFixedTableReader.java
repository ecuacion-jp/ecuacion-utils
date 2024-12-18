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
package jp.ecuacion.util.poi.read.string.reader;

import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.read.string.reader.internal.PoiStringTableReader;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Reads tables with known number of columns, known header labels 
 * and known start position of the table.
 * 
 * <p>The header line is required.
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in a line.</p>
 */
public abstract class PoiStringFixedTableReader extends PoiStringTableReader {

  /** 
   * Defines the header labels in the excel table. 
   * 
   * <p>It can be like {@code new String[] {"first name", "last name", "age"}}.</p>
   */
  protected abstract String[] getHeaderLabels();

  /**
   * Constructs a new instance. the obtained value 
   *     from an empty cell is {@code null}.
   * 
   * <p>In most cases {@code null} is recommended 
   *     because {@code Bean Validation} annotations (like {@code Max}) 
   *     returns valid for {@code null}, but invalid for {@code ""}.</p>
   */
  public PoiStringFixedTableReader() {
    super();
  }

  /**
   * Constructs a new instance with the obtained value from an empty cell.
   * 
   * @param noDataString the obtained value from an empty cell. {@code null} or {@code ""}.
   */
  public PoiStringFixedTableReader(NoDataString noDataString) {
    super(noDataString);
  }

  /**
   * Gets table data list in the form of {@code List<List<String>>}.
   * 
   * <p>header line is not included.</p>
   *
   * @see jp.ecuacion.util.poi.read.string.reader.internal.PoiStringTableReader
   */
  protected List<List<String>> getTableValues(String excelPath)
      throws AppException, EncryptedDocumentException, IOException {

    List<List<String>> rtnList = getTableValuesCommon(excelPath);

    // ヘッダ行のチェック
    List<String> headerList = rtnList.remove(0);
    for (int i = 0; i < headerList.size(); i++) {
      if (!headerList.get(i).equals(getHeaderLabels()[i])) {
        throw new BizLogicAppException("MSG_ERR_HEADER_TITLE_WRONG", getSheetName(),
            Integer.toString(i), headerList.get(i), getHeaderLabels()[i]);
      }
    }

    return rtnList;
  }

  @Override
  protected int getTableStartRowNumber(Sheet sheet) {
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

      if (value.equals(getHeaderLabels()[0])) {
        // iはプログラム上の行（ゼロ始まり）だが、getTableStartRowNumber()としては左上が(1, 1)として返すルールなので1をプラスして返す
        return i + 1;
      }
    }

    // ここまでくるということは、signStringがなかったということ。異常終了
    throw new RuntimeException(
        "シート「" + sheet.getSheetName() + "」に文字列「" + getHeaderLabels()[0] + "」がA列に存在しません。終了します。");
  }

  /**
   * Obtains table row size.
   * 
   * <p>固定のtableであれば、間に空行があるtableは極めて想像しにくいので、空行があればそこで読み込み終了の前提とする。 
   *     （@Overrideして固定値を設定することは一応許している）
   */
  @Override
  protected Integer getTableRowSize() {
    return null;
  }

  protected int getTableStartColumnNumber() {
    return 1;
  }

  protected int getTableStartColumnSize() {
    return getHeaderLabels().length;
  }
}
