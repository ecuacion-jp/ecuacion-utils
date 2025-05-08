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
package jp.ecuacion.util.poi.excel.enums;

/**
 * Has the selections which {@code java} gets when the value of a cell in an excel file is empty.
 * 
 * <p>The word {@code empty} usually means a string is either {@code null or ""}, 
 *     but {@code java empty string} or for short {@code empty string}, exactly means {@code ""}.
 *     <br>
 *     In this context hese two words are clearly distinguished so don't understand this wrong.
 *     
 * <p>{@code NULL} means {@code null}, {@code EMPTY_STRING} means {@code ""}.<br>
 * <b>{@code NULL} is recommended</b> 
 *     because usually values obtained from an excel file are validated 
 *     with {@code jakarta validation}, 
 *     and it consider {@code null} as valid, but {@code empty("")}  as invalid.</p>
 */
public enum NoDataString {
  
  /**
   * means {@code null}. <b>Recommended.</b>
   */
  NULL, 
  
  /**
   * means {@code empty string ("")}. <b>NOT Recommended.</b>
   */
  EMPTY_STRING;
}
