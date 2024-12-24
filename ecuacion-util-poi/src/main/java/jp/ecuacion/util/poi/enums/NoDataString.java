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
package jp.ecuacion.util.poi.enums;

/**
 * Has the selections which {@code java} gets when the value of a cell in an excel file is empty.
 * 
 * <p>{@code NULL} means null, {@code EMPTY} means "".<br>
 * <b>{@code NULL} is recommended</b> 
 *     because usually values obtained from an excel file are validated 
 *     with {@code bean validation}, 
 *     and it consider {@code null} as valid, but {@code empty("") as invalid.} </p>
 */
public enum NoDataString {
  
  /**
   * means {@code null}. <b>Recommended.</b>
   */
  NULL, 
  
  /**
   * means {@code empty ("")}. <b>NOT Recommended.</b>
   */
  EMPTY;
}
