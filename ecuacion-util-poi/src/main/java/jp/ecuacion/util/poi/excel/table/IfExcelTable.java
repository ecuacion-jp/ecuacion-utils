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
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;

/**
 * Provides the methods the extending interfaces use.
 * 
 * @param <T> The data type obtained from the excel table. 
 *     For example it would be {@code String} 
 *     if you want {@code String} data from each cell in the table.
 */
public interface IfExcelTable<T> {

  /**
   * Returns the excel sheet name the {@code TableReader} and the {@code TableWriter} access.
   * 
   * @return the sheet name of the excel file
   */
  @Nonnull
  public abstract String getSheetName();

  /**
   * Returns the value of the far left header cell to specify the position of the table.
   * 
   * <p>The method is called when {@code tableStartRowNumber} is {@code null}.<br>
   * See {@link ExcelTable#tableStartRowNumber}</p>
   * 
   * <p>When the table doesn't have a header, 
   *     an {@code exception} is thrown if {@code tableStartRowNumber} is {@code null}.<br>
   *     So always set non-null {@code tableStartRowNumber} value 
   *     when the table doesn't have a header.</p>
   * 
   * @return far left header label
   */
  @Nonnull
  public String getFarLeftHeaderLabel();

  /**
   * Validates the excel table header.
   * 
   * <p>If the table doesn't have a header, nothing needs to be done in this method.</p>
   *
   * @param headerData string header data 
   * @throws BizLogicAppException BizLogicAppException
   */
  public void validateHeader(@RequireNonnull List<List<String>> headerData)
      throws BizLogicAppException;

  /**
   * Is used to get the header label string from the argument cell data.
   * 
   * @param cellData data obtained from the cell
   * @return {@code String} value obtained from the {@code cellData}
   */
  @Nullable
  public String getStringValue(@Nullable T cellData);
  
}
