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
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.util.ObjectsUtil;

/**
 * Is a reader interface which treats one line header format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfFormatOneLineHeaderExcelTable<T>
    extends IfExcelTable<T> {

  /** 
   * Defines the header labels in the excel table. 
   * 
   * <p>It can be like {@code new String[] {"first name", "last name", "age"}}.</p>
   */
  @Nonnull
  public String[] getHeaderLabels();

  @Override
  public default String[][] getHeaderLabelData() {
    String[][] rtn = new String[1][];
    rtn[0] = getHeaderLabels();
    
    return rtn;
  }
  
  /**
   * Validates one line header.
   * 
   * <p>Pass a list with `size() == 0` 
   *     when it's a table with no header or nothing to validate.</p>
   * 
   * <p>See {@link IfExcelTable#validateHeaderData}.</p>
   * 
   * @param headerList headerList
   * @throws BizLogicAppException BizLogicAppException
   */
  public default void validateHeader(@RequireNonnull List<String> headerList)
      throws BizLogicAppException {
    List<List<String>> list = new ArrayList<>();
    list.add(headerList);
    
    validateHeaderData(list);
  }


  @Override
  @Nonnull
  public default String getFarLeftAndTopHeaderLabel() {

    String[] headerLabels = getHeaderLabels();
    ObjectsUtil.paramSizeNonZero(headerLabels);
    
    String farLeftHeaderLabel = headerLabels[0];
    return ObjectsUtil.paramRequireNonNull(farLeftHeaderLabel);
  }
}
