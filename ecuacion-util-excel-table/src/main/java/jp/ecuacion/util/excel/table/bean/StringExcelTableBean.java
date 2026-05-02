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

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.constant.EclibCoreConstants;
import jp.ecuacion.lib.core.logging.DetailLogger;
import org.jspecify.annotations.Nullable;

/**
 * Stores values obtained from excel tables with {@code StringFixedTableToBeanReader}.
 */
public abstract class StringExcelTableBean {

  private DetailLogger detailLog = new DetailLogger(this);

  /**
   * Is called after reading an excel file. 
   * 
   * <p>This is assumed to use to deserialize (structure) the line of data into objects, 
   *     and validate the inter-fields data.<br>
   *     Validations for each field are supposed to be done by bean vaildation.
   *     This method covers selective-requirement, or other inter-fields validations.</p>
   */
  public void afterReading() {

  }

  /**
   * Returns {@code String} array of field names in the bean.
   * 
   * <p>For example, the table in an excel file is this.</p>
   * <table border="1" style="border-collapse: collapse">
   * <tr>
   * <th>name</th>
   * <th>age</th>
   * <th>phone number</th>
   * </tr>
   * <tr>
   * <td>John</td>
   * <td>30</td>
   * <td>(+01)123456789</td>
   * </tr>
   * <tr>
   * <td>Ken</td>
   * <td>40</td>
   * <td>(+81)987654321</td>
   * </tr>
   * <caption>table 1</caption>
   * </table>
   * 
   * <p>If you want to read all data from table 1 and put into a bean, 
   *     you need to create a bean extends this class, 
   *     define fields {@code name, age, and phoneNumber}, 
   *     override this method and return the following array.</p>
   * 
   * <code>new String[] {"name", "age", "phoneNumber"}</code>
   * 
   * <p>the value in the first column (name) is put into 
   *     the field with the first element in the array (name).</p>
   *     
   * <p>If you don't need data in "age" column, you can set null like </p>
   * <code>new String[] {"name", null, "phoneNumber"}</code>
   */
  protected abstract String[] getFieldNameArray();

  /**
   * Constructs a new instance with the list of strings 
   *     which consists of data of a line from the excel table.
   *     
   * @param colList the list of strings which consists of data of a line from the excel table
   */
  public StringExcelTableBean(List<String> colList) {
    String[] fieldNameArray = getFieldNameArray();

    if (colList.size() != fieldNameArray.length) {
      throw new RuntimeException(
          "Number of elements in fieldNameArray and colList differ.\n" + "fieldNameArray ("
              + fieldNameArray.length + " elements) = " + Arrays.toString(getFieldNameArray())
              + ",\n" + "colList (" + colList.size() + " elements) = " + colList.toString());
    }

    try {
      detailLog.debug(EclibCoreConstants.PARTITION_LARGE);
      detailLog.debug("Setting values from excel file to bean started.");
      detailLog.debug("class name: " + this.getClass().getSimpleName());

      for (int i = 0; i < fieldNameArray.length; i++) {
        String fieldName = fieldNameArray[i];

        // null means this column is intentionally skipped (no corresponding field).
        if (fieldName == null) {
          continue;
        }

        // Walk up the class hierarchy to find the field, including inherited fields.
        Field field = null;
        Class<?> clazz = this.getClass();
        while (clazz != null) {
          try {
            field = clazz.getDeclaredField(fieldName);
            break;

          } catch (NoSuchFieldException ignored) {
            clazz = clazz.getSuperclass();

            if (clazz == null) {
              throw new RuntimeException("Trying to set a string value to the field in the bean, "
                  + "but the fieldName not found in the bean. \nbeanName: "
                  + this.getClass().getSimpleName() + ", fieldName: " + fieldName);
            }
          }
        }

        Objects.requireNonNull(field);

        field.setAccessible(true);
        field.set(this, colList.get(i));
      }

      detailLog.debug("Setting values from excel file to bean finished successfully.");
      detailLog.debug(EclibCoreConstants.PARTITION_LARGE);

    } catch (Exception ex) {
      throw new RuntimeException(ex);
    }
  }

  /** Returns {@code empty} if the argument value is null or returns the argument value. */
  protected String nullToEmpty(@Nullable String value) {
    return value == null ? "" : value;
  }

  /** Returns {@code null} if the argument value is empty or returns the argument value. */
  @Nullable
  protected String emptyToNull(@Nullable String value) {
    return value == null || value.equals("") ? null : value;
  }

  /** Returns {@code Integer} datatype of the argument string. */
  @Nullable
  protected Integer toInteger(@Nullable String value) {
    return value == null || value.equals("") ? null : Integer.valueOf(value);
  }

  /** Returns {@code Long} datatype of the argument string. */
  @Nullable
  protected Long toLong(@Nullable String value) {
    return value == null || value.equals("") ? null : Long.valueOf(value);
  }

  /** Returns {@code Float} datatype of the argument string. */
  @Nullable
  protected Float toFloat(@Nullable String value) {
    return value == null || value.equals("") ? null : Float.valueOf(value);
  }

  /** Returns {@code Double} datatype of the argument string. */
  @Nullable
  protected Double toDouble(@Nullable String value) {
    return value == null || value.equals("") ? null : Double.valueOf(value);
  }

  /** Returns {@code BigInteger} datatype of the argument string. */
  @Nullable
  protected BigInteger toBigInteger(@Nullable String value) {
    return value == null || value.equals("") ? null : new BigInteger(value);
  }

  /** Returns {@code BigDecimal} datatype of the argument string. */
  @Nullable
  protected BigDecimal toBigDecimal(@Nullable String value) {
    return value == null || value.equals("") ? null : new BigDecimal(value);
  }
}
