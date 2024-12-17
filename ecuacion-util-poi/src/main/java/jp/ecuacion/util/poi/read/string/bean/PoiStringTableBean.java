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

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.LogUtil;

public abstract class PoiStringTableBean {

  private DetailLogger detailLog = new DetailLogger(this);

  /** bean内の複数field間の整合性をチェック。 */
  public abstract void dataConsistencyCheck() throws AppException;
  
  /**
   * Excelからrowを読み込んだlist内の値を、指定の順に変数に設定する。 例えば、戻り値がnew String[] {"field1", null, "field2"} の場合、
   * field1 = list.get(0); field2 = list.get(2); のように設定されるイメージ。
   * listから値を取得しない列がある場合（excelの読み込み対象列にskipがある場合にこうなる）、nullで設定する。
   * 
   * <p>
   * 厳密には、constructorの引数に設定されるlistは、excelから読み込んだそのままのlistである必要はない。
   * 例えば、excel一覧上に分類項目と明細項目があり、読み込んだ結果分類と迷彩を別オブジェクトとして親子関係で保持する場合などは、
   * listをそれぞれのオブジェクトに必要な項目に絞り、getFieldNameArrary()もそれに応じた変数のみ設定することで読み込み可能。
   * ただ、skip箇所はnullと明示的に記載しexcelの列と整合がとれた記載の方がわかりやすいと思われる。
   * </p>
   */
  protected abstract String[] getFieldNameArray();

  /**
   * 通常は使用を想定されていないが、イレギュラーなconstructorを作成したい場合に使用。
   */
  public PoiStringTableBean() {}

  public PoiStringTableBean(List<String> colList) {
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

  /** nullを空文字に変換するためのutility method. */
  protected String nullToEmpty(String value) {
    return value == null ? "" : value;
  }

  /** 空文字をnullに変換するためのutility method. */
  protected String emptyToNull(String value) {
    return value.equals("") ? null : value;
  }

  /** getterでIntegerに変換するためのutility method. */
  protected Integer toInteger(String value) {
    return value == null || value.equals("") ? null : Integer.valueOf(value);
  }

  /** getterでLongに変換するためのutility method. */
  protected Long toLong(String value) {
    return value == null || value.equals("") ? null : Long.valueOf(value);
  }

  /** getterでFloatに変換するためのutility method. */
  protected Float toFloat(String value) {
    return value == null || value.equals("") ? null : Float.valueOf(value);
  }

  /** getterでDoubleに変換するためのutility method. */
  protected Double toDouble(String value) {
    return value == null || value.equals("") ? null : Double.valueOf(value);
  }

  /** getterでBigIntegerに変換するためのutility method. */
  protected BigInteger toBigInteger(String value) {
    return value == null || value.equals("") ? null : new BigInteger(value);
  }

  /** getterでBigDecimalに変換するためのutility method. */
  protected BigDecimal toBigDecimal(String value) {
    return value == null || value.equals("") ? null : new BigDecimal(value);
  }
}
