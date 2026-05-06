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
package jp.ecuacion.util.excel.table.bean;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Marks a field as mapped to an Excel column by header label.
 *
 * <p>When a {@link StringExcelTableBean} subclass annotates its fields with this annotation,
 *     overriding {@link StringExcelTableBean#getFieldNameArray()} is not required.<br>
 *     The {@link jp.ecuacion.util.excel.table.reader.concrete
 *     .StringOneLineHeaderExcelTableToBeanReader} matches each annotated field to the column
 *     whose header label equals {@link #value()}, regardless of column order in the Excel
 *     file.</p>
 *
 * <p>For single-row headers, pass one string:</p>
 * <pre>{@code
 * @ExcelColumn("name") String name;
 * }</pre>
 *
 * <p>For multi-row headers, pass one string per header row (top to bottom).
 *     If the column has the same label in every header row (vertically merged),
 *     a single string can also be used:</p>
 * <pre>{@code
 * // 2-row header: group row + column row
 * @ExcelColumn({"個人情報", "名前"}) String name;
 *
 * // Vertically merged column – same value in all header rows
 * @ExcelColumn("#") Integer rowNumber;
 * }</pre>
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

  /**
   * The header label(s) of the corresponding Excel column, one entry per header row
   *     from top to bottom.
   *
   * <p>A single-element array (or a plain string) matches any column whose header labels
   *     are all equal to that element (covers both single-row and vertically-merged columns).</p>
   *
   * @return header label(s)
   */
  String[] value();
}
