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

/**
 * Is a reader interface which treats free format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfFormatFreeExcelTable<T> extends IfExcelTable<T> {

  @Override
  public default void validateHeaderData(@Nullable List<List<T>> headerData) {
    // no validations for the argument excel data.
  }

  @Override
  public default int getNumberOfHeaderLines() {
    return 0;
  }
  
  @Override
  @Nonnull
  public default String[][] getHeaderLabelData() {
    throw new NullPointerException();
  }

  @Override
  @Nonnull
  public default String getFarLeftAndTopHeaderLabel() {
    throw new NullPointerException();
  }
}
