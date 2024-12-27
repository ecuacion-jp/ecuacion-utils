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

/**
 * Provides excel-related {@code apache POI} utility classes.
 */
module jp.ecuacion.util.poi {
  exports jp.ecuacion.util.poi.excel.table.reader.string.internal;
  exports jp.ecuacion.util.poi.excel.enums;
  exports jp.ecuacion.util.poi.excel.util;
  exports jp.ecuacion.util.poi.excel.table.reader.cell;
  exports jp.ecuacion.util.poi.excel.table.reader.core;
  exports jp.ecuacion.util.poi.excel.table.reader.string.bean;
  exports jp.ecuacion.util.poi.excel.table.reader.string;
  exports jp.ecuacion.util.poi.excel.table.writer.cell;
  exports jp.ecuacion.util.poi.excel.table.writer.core;

  requires jakarta.annotation;
  requires jakarta.validation;
  requires jp.ecuacion.lib.core;
  requires org.apache.commons.lang3;
  requires transitive org.apache.poi.poi;
  requires org.apache.poi.ooxml;
}
