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
package jp.ecuacion.util.poi.read.core.reader;

import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;

/**
 * Features the format of the table.
 * 
 * <p></p>
 * 
 * @param <T> Object obtained from the excel data. 
 *     It would be {@code String} if you get {@code String} data from each cell in the table.
 */
public interface IfAbstractFormatTableReader<T> {

  /**
   * Returns the value of the left cell in the header line.
   * 
   * @return header label
   */
  public String getHeaderLabelToDecideTableStartRowNumber();
  
  /**
   * Returns a list of lines ({@literal lines} is also a list of string cell value).
   *
   * @throws BizLogicAppException BizLogicAppException
   *
   * @see jp.ecuacion.util.poi.read.string.reader.internal.StringTableReader
   */
  public void validateAndUpdate(List<List<T>> excelTableData)
      throws BizLogicAppException;

}
