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
package jp.ecuacion.util.poi.excel.table.reader.core;

import jakarta.annotation.Nonnull;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.IfOneLineHeaderFormatExcelTable;

/**
 * Is a reader which treats one line header format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfOneLineHeaderFormatExcelTableReader<T>
    extends IfOneLineHeaderFormatExcelTable<T>, IfExcelTableReader<T> {

  @Override
  public default List<List<String>> updateAndGetHeaderList(@Nonnull List<List<T>> excelData) {
    List<String> list = excelData.remove(0).stream().map(el -> getStringValue(el)).toList();
    
    List<List<String>> rtnList = new ArrayList<>();
    rtnList.add(list);
    
    return rtnList;
  }
}
