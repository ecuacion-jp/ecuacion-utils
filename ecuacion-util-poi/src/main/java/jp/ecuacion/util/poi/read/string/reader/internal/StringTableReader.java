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

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.read.core.reader.TableReader;
import jp.ecuacion.util.poi.util.PoiReadUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

public abstract class StringTableReader extends TableReader<String> {
  protected PoiReadUtil readUtil;

  public StringTableReader(@RequireNonnull String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);

    readUtil = new PoiReadUtil();
  }

  public StringTableReader(@RequireNonnull String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize,
      @Nonnull NoDataString noDataString) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);

    readUtil = new PoiReadUtil(noDataString);
  }

  @Override
  public String getCellData(Cell cell) {
    return readUtil.getStringFromCell(cell);
  }

  @Override
  public boolean isCellDataEmpty(@Nullable String cellData) {
    return StringUtils.isEmpty(cellData);
  }
}
