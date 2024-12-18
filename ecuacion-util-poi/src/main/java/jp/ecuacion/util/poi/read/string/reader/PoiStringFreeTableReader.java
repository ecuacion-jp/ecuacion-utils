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
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.read.string.reader.internal.PoiStringTableReader;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Reads tables with unknown number of columns, unknown header labels 
 * and unknown start position of the table.
 * 
 * <p>The header line is not necessary. 
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in one line.</p>
 */
public abstract class PoiStringFreeTableReader extends PoiStringTableReader {

  /* getTableStartRowNumber() にて返される値。 */
  private int tableStartRowNumber;

  /* getTableRowSize() にて返される値。 */
  private Integer tableRowSize;

  /* getTableStartColumnNumber() にて返される値。 */
  private int tableStartColumnNumber;

  /* getTableStartColumnSize() にて返される値。 */
  private int tableColumnSize;

  /**
   * Constructs a new instance.
   * 
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableColumnSize tableColumnSize
   */
  public PoiStringFreeTableReader(int tableStartRowNumber, int tableStartColumnNumber,
      int tableColumnSize) {
    this(tableStartRowNumber, tableStartColumnNumber, tableColumnSize, NoDataString.NULL);
  }

  /**
   * Constructs a new instance.
   * 
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableColumnSize tableColumnSize
   * @param noDataString noDataString
   */
  public PoiStringFreeTableReader(int tableStartRowNumber, int tableStartColumnNumber,
      int tableColumnSize, NoDataString noDataString) {
    this(tableStartRowNumber, null, tableStartColumnNumber, tableColumnSize, noDataString);
  }

  /**
   * Constructs a new instance.
   * 
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableRowSize tableRowSize
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableColumnSize tableColumnSize
   */
  public PoiStringFreeTableReader(int tableStartRowNumber, Integer tableRowSize,
      int tableStartColumnNumber, int tableColumnSize) {
    this(tableStartRowNumber, tableRowSize, tableStartColumnNumber, tableColumnSize,
        NoDataString.NULL);
  }

  /**
   * Constructs a new instance.
   * 
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableRowSize tableRowSize
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableColumnSize tableColumnSize
   * @param noDataString noDataString
   */
  public PoiStringFreeTableReader(int tableStartRowNumber, Integer tableRowSize,
      int tableStartColumnNumber, int tableColumnSize, NoDataString noDataString) {
    super(noDataString);

    this.tableStartRowNumber = tableStartRowNumber;
    this.tableRowSize = tableRowSize;
    this.tableStartColumnNumber = tableStartColumnNumber;
    this.tableColumnSize = tableColumnSize;
  }

  @Override
  protected int getTableStartRowNumber(Sheet sheet) {
    return tableStartRowNumber;
  }

  @Override
  protected Integer getTableRowSize() {
    return tableRowSize;
  }

  @Override
  protected int getTableStartColumnNumber() {
    return tableStartColumnNumber;
  }

  @Override
  protected int getTableStartColumnSize() {
    return tableColumnSize;
  }

  /**
   * Returns a list of lines ({@literal lines} is also a list of string cell value).
   * 
   * @return a list of lines, may contain header line if the original table has the header.
   *
   * @see jp.ecuacion.util.poi.read.string.reader.internal.PoiStringTableReader
   */
  protected List<List<String>> getTableValues(String excelPath)
      throws EncryptedDocumentException, AppException, IOException {
    return getTableValuesCommon(excelPath);
  }
}
