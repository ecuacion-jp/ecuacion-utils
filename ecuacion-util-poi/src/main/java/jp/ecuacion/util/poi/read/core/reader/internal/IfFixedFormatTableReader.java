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
package jp.ecuacion.util.poi.read.core.reader.internal;

import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.read.core.reader.IfAbstractFormatTableReader;

public interface IfFixedFormatTableReader<T> extends IfAbstractFormatTableReader<T> {

  public abstract String getSheetName();
  
  /** 
   * Defines the header labels in the excel table. 
   * 
   * <p>It can be like {@code new String[] {"first name", "last name", "age"}}.</p>
   */
  public String[] getHeaderLabels();

  public default void validateAndUpdate(List<List<T>> tableData)
      throws BizLogicAppException {
    // ヘッダ行のチェック。同時にヘッダ行はexcelTableDataListからremoveしておき、returnするデータには含めない
    List<String> headerList =
        tableData.remove(0).stream().map(el -> getCellStringValue(el)).toList();
    
    for (int i = 0; i < headerList.size(); i++) {
      if (!headerList.get(i).equals(getHeaderLabels()[i])) {
        throw new BizLogicAppException("MSG_ERR_HEADER_TITLE_WRONG", getSheetName(),
            Integer.toString(i), headerList.get(i), getHeaderLabels()[i]);
      }
    }
  }

  /**
   * Is Used to get Header Label String from the argument cell data.
   */
  public String getCellStringValue(T cellData);


  @Override
  public default String getHeaderLabelToDecideTableStartRowNumber() {
    return getHeaderLabels()[0];
  }
}
