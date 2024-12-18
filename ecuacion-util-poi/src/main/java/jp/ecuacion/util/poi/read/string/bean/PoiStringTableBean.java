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
package jp.ecuacion.util.poi.read.string.bean;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.LogUtil;

/**
 * Stores values obtained from excel tables with {@code PoiStringFixedTableToBeanReader}.
 */
public abstract class PoiStringTableBean {

  private DetailLogger detailLog = new DetailLogger(this);

  /**
   * Validates the inter-fields data.
   * 
   * <p>Validations for each field needs to be done by bean vaildation.
   *     This method covers selective-requirement, or other inter-fields validations.
   */
  public abstract void dataConsistencyCheck() throws AppException;

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
  @Nonnull
  protected abstract String[] getFieldNameArray();

  // /**
  // * 通常は使用を想定されていないが、イレギュラーなconstructorを作成したい場合に使用。
  // */
  // public PoiStringTableBean() {}

  /**
   * Constructs a new instance with the list of strings 
   *     which consists of data of a line from the excel table.
   *     
   * @param colList the list of strings which consists of data of a line from the excel table
   */
  public PoiStringTableBean(@Nonnull List<String> colList) {
    String[] fieldNameArray = getFieldNameArray();

    // colListの件数とfieldNameArraryの件数が異なる場合はエラー
    if (colList.size() != fieldNameArray.length) {
      throw new RuntimeException(
          "Number of elements in fieldNameArray and colList differ.\n" + "fieldNameArray ("
              + fieldNameArray.length + " elements) = " + Arrays.toString(getFieldNameArray())
              + ",\n" + "colList (" + colList.size() + " elements) = " + colList.toString());
    }

    try {
      detailLog.debug(LogUtil.PARTITION_LARGE);
      detailLog.debug("excelファイルから読み込んだ値のbeanへの設定開始");
      detailLog.debug("class名：" + this.getClass().getSimpleName());

      for (int i = 0; i < fieldNameArray.length; i++) {
        String fieldName = fieldNameArray[i];

        // nullの場合は、excelの対象列から値を取得しておらず、設定する変数もないという意味なのでskip
        if (fieldName == null) {
          continue;
        }

        // 親クラスのfieldも取得できるよう、親クラスを際気的に検索してfieldを取得
        Field field = null;
        Class<?> clazz = this.getClass();
        while (clazz != null) {
          try {
            field = clazz.getDeclaredField(fieldName);
            break;

          } catch (NoSuchFieldException e) {
            // 親のクラスのfieldを探す
            clazz = clazz.getSuperclass();

            // clazz == nullの場合は、一番親まで遡ったがfieldが存在しない、つまりfieldNameの指定が間違い
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

      detailLog.debug("（excelファイルから読み込んだ値のbeanへの設定正常終了）");
      detailLog.debug(LogUtil.PARTITION_LARGE);

    } catch (Exception ex) {
      throw new RuntimeException(ex);
    }
  }

  /** Returns {@code empty} if the argument value is null or returns the argument value. */
  @Nonnull
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
