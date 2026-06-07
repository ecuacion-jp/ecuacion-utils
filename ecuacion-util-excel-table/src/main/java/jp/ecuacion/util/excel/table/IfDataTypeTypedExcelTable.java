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
package jp.ecuacion.util.excel.table;

import jp.ecuacion.util.excel.exception.ExcelTableException;
import org.jspecify.annotations.Nullable;

/**
 * Provides the excel table interface
 *     with object type obtained from the excel data being a native Java type.
 *
 * <p>Numeric cells return {@link Double}, date-formatted cells return
 *     {@link java.time.LocalDate} or {@link java.time.LocalDateTime},
 *     string cells return {@link String}, and boolean cells return {@link Boolean}.</p>
 */
public interface IfDataTypeTypedExcelTable extends IfExcelTable<Object> {

  @Override
  @Nullable
  public default String getStringValue(@Nullable Object cellData) throws ExcelTableException {
    return cellData == null ? null : cellData.toString();
  }
}
