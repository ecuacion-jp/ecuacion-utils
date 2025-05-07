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
package jp.ecuacion.util.poi.excel.table.reader;

import jakarta.annotation.Nonnull;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;

/**
 * Is a reader which treats one line header format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfFormatOneLineHeaderExcelTableReader<T>
    extends IfFormatOneLineHeaderExcelTable<T>, IfExcelTableReader<T> {

  @Override
  public default List<List<String>> updateAndGetHeaderData(@Nonnull List<List<T>> excelData)
      throws ExcelAppException {
    List<List<String>> list = new ArrayList<>();
    
    if (excelData.size() == 0) {
      return list;
    }
    
    List<T> headerLineList = excelData.remove(0);
    List<String> strList = new ArrayList<>();
    for (T el : headerLineList) {
      strList.add(getStringValue(el));
    }
    
    list.add(strList);
    
    return list;
  }
}
