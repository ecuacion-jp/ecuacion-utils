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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
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
   * Returns {@code String} array of field names that correspond to Excel columns, in the order
   *     the values will be received from the reader.
   *
   * <p>The default implementation scans the class hierarchy for fields annotated with
   *     {@link ExcelColumn} and returns their field names. When {@link ExcelColumn} annotations
   *     are present, overriding this method is not required — the reader handles column-order
   *     matching by header label automatically.</p>
   *
   * <p>Override this method when not using {@link ExcelColumn} annotations.</p>
   *
   * <p>Example (manual override):</p>
   * <table border="1" style="border-collapse: collapse">
   * <caption>table 1</caption>
   * <tr><th>name</th><th>age</th><th>phone number</th></tr>
   * <tr><td>John</td><td>30</td><td>(+01)123456789</td></tr>
   * </table>
   *
   * <pre>{@code
   * @Override
   * protected String[] getFieldNameArray() {
   *   return new String[] {"name", "age", "phoneNumber"};
   * }
   * }</pre>
   *
   * <p>Set {@code null} to skip a column:
   *     {@code new String[] {"name", null, "phoneNumber"}}</p>
   *
   * @throws RuntimeException if no {@link ExcelColumn} annotations are found and this method
   *     is not overridden
   */
  protected String[] getFieldNameArray() {
    List<String> fieldNames = new ArrayList<>();
    List<Class<?>> hierarchy = new ArrayList<>();
    Class<?> clazz = this.getClass();
    while (clazz != null && clazz != StringExcelTableBean.class) {
      hierarchy.add(0, clazz);
      clazz = clazz.getSuperclass();
    }
    for (Class<?> c : hierarchy) {
      for (Field f : c.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          fieldNames.add(f.getName());
        }
      }
    }
    if (fieldNames.isEmpty()) {
      throw new RuntimeException("No @ExcelColumn annotations found in "
          + this.getClass().getSimpleName()
          + ". Either annotate fields with @ExcelColumn or override getFieldNameArray().");
    }
    return fieldNames.toArray(new String[0]);
  }

  /**
   * Constructs a new instance with the list of strings
   *     which consists of data of a line from the excel table.
   *
   * <p>Field types are automatically detected via reflection. Fields declared as
   *     {@code Integer}, {@code LocalDate}, {@code BigDecimal}, and other supported types
   *     are converted from the string value read from Excel.
   *     See {@link #convertToFieldType} for supported types.</p>
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
        field.set(this, convertToFieldType(field.getType(), colList.get(i), fieldName));
      }

      detailLog.debug("Setting values from excel file to bean finished successfully.");
      detailLog.debug(EclibCoreConstants.PARTITION_LARGE);

    } catch (RuntimeException ex) {
      throw ex;
    } catch (Exception ex) {
      throw new RuntimeException(ex);
    }
  }

  /**
   * Returns the {@code DateTimeFormatter} used when converting string values to date/time types.
   *
   * <p>Defaults to {@code DateTimeFormatter.ISO_LOCAL_DATE} ({@code yyyy-MM-dd}).
   *     Override this method in a subclass to use a different format.</p>
   *
   * @return DateTimeFormatter for date conversion
   */
  protected DateTimeFormatter getDateTimeFormatter() {
    return DateTimeFormatter.ISO_LOCAL_DATE;
  }

  /**
   * Converts a string value from an Excel cell to the declared type of the target field.
   *
   * <p>Supported types: {@code String}, {@code Integer}/{@code int},
   *     {@code Long}/{@code long}, {@code Short}/{@code short},
   *     {@code Float}/{@code float}, {@code Double}/{@code double},
   *     {@code BigDecimal}, {@code BigInteger}, {@code Boolean}/{@code boolean},
   *     {@code LocalDate}, {@code LocalDateTime}, {@code LocalTime}.</p>
   *
   * <p>Returns {@code null} for {@code null} or empty string input,
   *     except for {@code String} fields which retain the value as-is.</p>
   *
   * @param fieldType the declared type of the target field
   * @param value the string value from the Excel cell, may be {@code null}
   * @param fieldName the field name, used for error messages
   * @return the converted value, or {@code null} for empty input on non-String types
   * @throws RuntimeException if conversion fails (e.g. non-numeric string for Integer field)
   */
  private @Nullable Object convertToFieldType(Class<?> fieldType,
      @Nullable String value, String fieldName) {
    boolean isEmpty = value == null || value.isEmpty();

    if (fieldType == String.class) {
      return value;
    }
    if (isEmpty) {
      return null;
    }

    try {
      if (fieldType == Integer.class || fieldType == int.class) {
        return Integer.valueOf(value);
      }
      if (fieldType == Long.class || fieldType == long.class) {
        return Long.valueOf(value);
      }
      if (fieldType == Short.class || fieldType == short.class) {
        return Short.valueOf(value);
      }
      if (fieldType == Float.class || fieldType == float.class) {
        return Float.valueOf(value);
      }
      if (fieldType == Double.class || fieldType == double.class) {
        return Double.valueOf(value);
      }
      if (fieldType == BigDecimal.class) {
        return new BigDecimal(value);
      }
      if (fieldType == BigInteger.class) {
        return new BigInteger(value);
      }
      if (fieldType == Boolean.class || fieldType == boolean.class) {
        return Boolean.valueOf(value);
      }
      if (fieldType == LocalDate.class) {
        return LocalDate.parse(value, getDateTimeFormatter());
      }
      if (fieldType == LocalDateTime.class) {
        return LocalDateTime.parse(value);
      }
      if (fieldType == LocalTime.class) {
        return LocalTime.parse(value);
      }
    } catch (Exception ex) {
      throw new RuntimeException("Failed to convert value '" + value + "' to type "
          + fieldType.getSimpleName() + " for field '" + fieldName + "'.", ex);
    }

    // Unknown type: attempt direct assignment (will fail at field.set if incompatible).
    return value;
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
