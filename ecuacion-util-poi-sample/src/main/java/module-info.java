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
module jp.ecuacion.util.poi.sample {
  exports jp.ecuacion.util.poi.sample.readstringfromcell;
  exports jp.ecuacion.util.poi.sample.copytable;

  requires jp.ecuacion.lib.core;
  requires jp.ecuacion.util.poi;
  requires org.apache.poi.poi;
  requires org.slf4j;
  
  opens jp.ecuacion.util.poi.sample.readstringfromcell;
}
