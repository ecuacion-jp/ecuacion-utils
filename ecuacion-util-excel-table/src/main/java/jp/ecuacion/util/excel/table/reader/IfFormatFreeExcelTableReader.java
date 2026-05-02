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
package jp.ecuacion.util.excel.table.reader;

import java.util.List;
import jp.ecuacion.util.excel.table.IfExcelTable;
import jp.ecuacion.util.excel.table.IfFormatFreeExcelTable;
import org.jspecify.annotations.Nullable;

/**
 * Is a reader which treats free format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfFormatFreeExcelTableReader<T>
    extends IfFormatFreeExcelTable<T>, IfExcelTableReader<T> {

  @Override
  public default @Nullable List<List<String>> updateAndGetHeaderData(List<List<T>> rtnData) {
    return null;
  }

  @Override
  public default void validateHeaderData(@Nullable List<List<T>> headerData) {
    // no validations for the argument excel data.
  }
}
