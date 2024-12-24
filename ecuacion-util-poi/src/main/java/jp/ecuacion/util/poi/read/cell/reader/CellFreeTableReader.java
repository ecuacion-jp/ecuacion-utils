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
package jp.ecuacion.util.poi.read.cell.reader;

import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.read.cell.reader.internal.CellTableReader;
import jp.ecuacion.util.poi.read.core.reader.internal.IfFreeFormatTableReader;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 
 */
public abstract class CellFreeTableReader extends CellTableReader
    implements IfFreeFormatTableReader<Cell> {

  /**
   * 
   * @param tableStartRowNumber
   * @param tableStartColumnNumber
   * @param tableRowSize
   * @param tableColumnSize
   */
  public CellFreeTableReader(@RequireNonnull String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }
}
