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
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.excel.table.reader.core.ExcelTableReader;

/**
 * features the format of the table.
 * 
 * @param <T> Object type obtained from the excel data. 
 *     For example, it would be {@code String} 
 *     if you want {@code String} data from each cell in the table.
 */
public interface IfExcelTable<T> {

  /**
   * Returns the excel sheet name the {@code TableReader} reads.
   * 
   * @return the sheet name
   */
  @Nonnull
  public abstract String getSheetName();

  /**
   * Returns the value of the far left header cell to specify the position of the table.
   * 
   * <p>Return value is used when {@code tableStartRowNumber} is {@code null}.<br>
   * See {@link ExcelTableReader#ExcelTableReader(String, Integer, int, Integer, Integer)}</p>
   * 
   * @return header label, may be {@code null} when the table don't have a header. 
   *     In that case, {@code Exception} is thrown if {@code tableStartRowNumber} is {@code null}.
   */
  @Nullable
  public String getFarLeftHeaderLabel();

  /**
   * Validate and Update the argument list, which is the table data obtained from an excel file.
   * 
   * <p>When the table you want to read is in a fixed format, 
   * this will validate header labels and remove the header line from the data because it's obvious
   *     and reduce the task to remove it by each caller methods.
   *
   * @throws BizLogicAppException BizLogicAppException
   */
  public void validateHeader(@Nullable List<List<String>> headerData)
      throws BizLogicAppException;

  /**
   * Is Used to get Header Label String from the argument cell data.
   */
  @Nullable
  public String getStringValue(@Nullable T cellData);
}
