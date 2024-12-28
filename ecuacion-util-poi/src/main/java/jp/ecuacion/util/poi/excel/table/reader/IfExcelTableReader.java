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
import jakarta.annotation.Nullable;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Provides the excel table reader methods.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfExcelTableReader<T> extends IfExcelTable<T> {

  /**
   * Returns an instance of {@code ExcelReadUtil}.
   * 
   * @return {@code ExcelReadUtil} instance
   */
  public ExcelReadUtil getExcelReadUtil();
  
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

  /**
   * Returns the obtained value from the cell.
   * 
   * <p>If you want to get {@code String} value from the cell, 
   *     it returns the {@code String} value.</p>
   * 
   * @param cell cell, may be null.
   * @return the obtained value from the cell
   */
  public @Nullable T getCellData(@RequireNonnull Cell cell);

  /**
   * Returns whether the value of the cell is empty.
   * 
   * @param cellData cellData
   * @return whether the valule of the cell is empty.
   */
  public boolean isCellDataEmpty(@Nullable T cellData);
}