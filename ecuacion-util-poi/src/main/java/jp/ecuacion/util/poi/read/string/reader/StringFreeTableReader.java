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

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.read.core.reader.internal.IfFreeFormatTableReader;
import jp.ecuacion.util.poi.read.string.reader.internal.StringTableReader;

/**
 * Reads a table with unknown number of columns, unknown whether it have a header line,
 * unknown header labels if it has a header line.
 * 
 * <p>The header line is not necessary. 
 *     This class reads the table at the designated position and designated lines and columns.<br>
 *     Finish reading if all the columns are empty in one line.</p>
 */
public class StringFreeTableReader extends StringTableReader
    implements IfFreeFormatTableReader<String> {

  /**
   * Constructs a new instance.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link jp.ecuacion.util.poi.read.core.reader.TableReader#TableReader(
   *     String, Integer, int, Integer, Integer)}.</p>
   *     
   * @param sheetName the sheet name of the excel file
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableRowSize tableRowSize
   * @param tableColumnSize tableRowSize
   */
  public StringFreeTableReader(@RequireNonnull String sheetName,
      @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, @Nullable Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }

  /**
   * Constructs a new instance.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber}, 
   *     {@code tableStartColumnNumber}, {@code tableRowSize} and {@code tableColumnSize},
   *     see {@link jp.ecuacion.util.poi.read.core.reader.TableReader#TableReader(
   *     String, Integer, int, Integer, Integer)}.</p>
   *     
   * @param sheetName the sheet name of the excel file
   * @param tableStartRowNumber tableStartRowNumber
   * @param tableStartColumnNumber tableStartColumnNumber
   * @param tableRowSize tableRowSize
   * @param tableColumnSize tableRowSize
   * @param noDataString noDataString
   */
  public StringFreeTableReader(@RequireNonnull String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize,
      @Nonnull NoDataString noDataString) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize,
        noDataString);

  }
}
