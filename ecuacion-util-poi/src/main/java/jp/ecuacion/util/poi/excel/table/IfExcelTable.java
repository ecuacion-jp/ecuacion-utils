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
import jp.ecuacion.lib.core.util.ObjectsUtil;

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
  public String getSheetName();

  /**
   * Returns the value of the far left and top header cell to specify the position of the table.
   * 
   * <p>The method is called when {@code tableStartRowNumber} is {@code null}.<br>
   * See {@link ExcelTable#tableStartRowNumber}</p>
   * 
   * <p>When the table doesn't have a header and {@code tableStartRowNumber} is {@code null},
   *     an {@code exception} is thrown.<br>
   *     So always set non-null {@code tableStartRowNumber} value 
   *     when the table doesn't have a header.</p>
   * 
   * @return far left and top header label<br>
   *     "top" means the upper side of the header line when the table has multiple header lines.
   */
  @Nonnull
  public String getFarLeftAndTopHeaderLabel();
  
  /**
   * Returns an array of header label strings.
   * 
   * <p>The data type of the return is {@code String[][]} 
   *     because table header can be multiple lines.</p>
   * 
   * @return table header label strings
   */
  @Nonnull
  public String[][] getHeaderLabelData();

  /**
   * Validates the excel table header.
   * 
   * @param headerData string header data<br>
   *     The data type is {@code List<List<String>> headerData} 
   *     because the header with multiple lines may exist.<br>
   *     Pass a list with `size() == 0` 
   *     when it's a table with no header or nothing to validate.
   * @throws BizLogicAppException BizLogicAppException
   */
  public default void validateHeaderData(@RequireNonnull List<List<String>> headerData)
      throws BizLogicAppException {
    
    for (int i = 0; i < ObjectsUtil.paramRequireNonNull(headerData).size(); i++) {
      List<String> headerList = headerData.get(i);
      String[] headerLabels = getHeaderLabelData()[i];
      
      for (int j = 0; j < headerList.size(); j++) {
        if (!headerList.get(j).equals(headerLabels[j])) {
          int positionFromUser = j + 1;
          throw new BizLogicAppException("MSG_ERR_HEADER_TITLE_WRONG", getSheetName(),
              Integer.toString(positionFromUser), headerList.get(j), headerLabels[j]);
        }
      }
    }
  }

  /**
   * Is used to get the header label string from the argument cell data.
   * 
   * @param cellData data obtained from the cell
   * @return {@code String} value obtained from the {@code cellData}
   */
  @Nullable
  public String getStringValue(@Nullable T cellData);
  
}
