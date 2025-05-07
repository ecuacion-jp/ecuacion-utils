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
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;

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
   * Returns the number of header lines.
   * 
   * @return the number of header lines
   */
  public int getNumberOfHeaderLines();

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
   * Stores the boolean value which indicates whether {@code validateHeaderData} ignores
   *     additional header columns.
   * 
   * @param value boolean
   */
  public IfExcelTable<T> ignoresAdditionalColumnsOfHeaderData(boolean value);

  /**
   * Obtains the boolean value which indicates whether {@code validateHeaderData} ignores
   *     additional header columns.
   * 
   * @return boolean
   */
  public boolean ignoresAdditionalColumnsOfHeaderData();

  /**
   * Is used to get the header label string from the argument cell data.
   * 
   * @param cellData data obtained from the cell
   * @return {@code String} value obtained from the {@code cellData}
   * @throws ExcelAppException ExcelAppException
   */
  @Nullable
  public String getStringValue(@Nullable T cellData) throws ExcelAppException;

  /**
   * Decides whether header is top (normal table) or left. 
   * {@code true} means headers are at the left.
   * 
   * @param value boolean
   * @return {@code IfExcelTable<T>}
   */
  public IfExcelTable<T> isVerticalAndHorizontalOpposite(boolean value);

  /**
   * Obtains whether header is top (normal table) or left. 
   * {@code true} means headers are at the left.
   */
  public boolean isVerticalAndHorizontalOpposite();
}
