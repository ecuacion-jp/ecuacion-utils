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
  public default void validateHeader(@Nullable List<List<String>> headerData)
      throws BizLogicAppException {
    
    // headerData is nonnull since headerData is in the format of one line header.
    List<String> headerList = ObjectsUtil.paramRequireNonNull(headerData).get(0);
    
    for (int i = 0; i < headerList.size(); i++) {
      if (!headerList.get(i).equals(getHeaderLabels()[i])) {
        throw new BizLogicAppException("MSG_ERR_HEADER_TITLE_WRONG", getSheetName(),
            Integer.toString(i), headerList.get(i), getHeaderLabels()[i]);
      }
    }
  }


  @Override
  @Nonnull
  public default String getFarLeftHeaderLabel() {

    String[] headerLabels = getHeaderLabels();
    ObjectsUtil.paramSizeNonZero(headerLabels);
    
    String farLeftHeaderLabel = headerLabels[0];
    return ObjectsUtil.paramRequireNonNull(farLeftHeaderLabel);
  }
}
