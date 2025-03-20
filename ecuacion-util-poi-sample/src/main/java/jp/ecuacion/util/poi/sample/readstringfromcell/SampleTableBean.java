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
package jp.ecuacion.util.poi.sample.readstringfromcell;

import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.util.poi.excel.table.bean.StringExcelTableBean;

public class SampleTableBean extends StringExcelTableBean {

  private String id;
  private String name;
  private String dateOfBirth;
  private String age;
  private String nationality;

  public SampleTableBean(@RequireNonnull List<String> colList) {
    super(colList);
  }

  @Override
  protected String[] getFieldNameArray() {
    return new String[] {"id", "name", "dateOfBirth", "age", "nationality"};
  }

  public String getId() {
    return id;
  }

  public String getName() {
    return name;
  }

  public String getDateOfBirth() {
    return dateOfBirth;
  }

  public String getAge() {
    return age;
  }

  public String getNationality() {
    return nationality;
  }

  public String toString() {
    return id + ", " + name + ", " + dateOfBirth.replaceAll("\\\\", "") + ", " + age + ", "
        + nationality;
  }
}
