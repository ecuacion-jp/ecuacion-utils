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
import jakarta.annotation.Nullable;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;

/**
 * features the format of the table.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfExcelTableReader<T> extends IfExcelTable<T> {

  /**
   * Updates excel data to treat it easily, like remove its header line, 
   *     and returns the header list.
   * 
   * <p>Considering various patterns of headers, return type ls {@code List<List<String>>}.</p>
   * 
   * @param tableData table data
   * @return header data
   */
  @Nullable
  public List<List<String>> updateAndGetHeaderList(@Nonnull List<List<T>> tableData);
}
