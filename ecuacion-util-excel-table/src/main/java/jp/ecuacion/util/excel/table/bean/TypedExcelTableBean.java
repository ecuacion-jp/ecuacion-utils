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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.constant.EclibCoreConstants;
import jp.ecuacion.lib.core.logging.DetailLogger;
import org.jspecify.annotations.Nullable;

/**
 * Stores values obtained from Excel tables with typed readers.
 *
 * <p>Unlike {@link StringExcelTableBean}, this class accepts native Java types directly:
 *     {@link Double} for numeric cells, {@link LocalDate} or {@link LocalDateTime} for
 *     date-formatted cells, {@link String} for string cells, and {@link Boolean} for boolean
 *     cells.</p>
 *
 * <p>When a numeric cell value ({@link Double}) is assigned to an integer-type field
 *     ({@code Integer}, {@code Long}, {@code Short}, {@code BigInteger}),
 *     the value is rounded via {@link Math#round(double)}, matching Excel's display behavior
 *     where integer cells store the value as a floating-point number internally.</p>
 *
 * <p>Annotate fields with {@link ExcelColumn} to let the reader match columns by header label
 *     regardless of column order.</p>
 */
public abstract class TypedExcelTableBean {

  private DetailLogger detailLog = new DetailLogger(this);

  /**
   * Called after reading an Excel row.
   *
   * <p>Override to implement cross-field validation or additional initialization.</p>
   */
  public void afterReading() {}

  /**
   * Returns the field names corresponding to Excel columns, in the order values are received.
   *
   * <p>The default implementation scans the class hierarchy for fields annotated with
   *     {@link ExcelColumn} and returns their names. Override when not using annotations.</p>
   *
   * @throws RuntimeException if no {@link ExcelColumn} annotations are found and this method
   *     is not overridden
   */
  protected @Nullable String[] getFieldNameArray() {
    List<String> fieldNames = new ArrayList<>();
    List<Class<?>> hierarchy = new ArrayList<>();
    Class<?> clazz = this.getClass();
    while (clazz != null && clazz != TypedExcelTableBean.class) {
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
   * Constructs a new instance from one Excel data row.
   *
   * <p>Each element in {@code colList} is a native Java value produced by a typed reader:
   *     {@link Double}, {@link LocalDate}, {@link LocalDateTime}, {@link String},
   *     {@link Boolean}, or {@code null} for empty cells.</p>
   *
   * @param colList typed values from one Excel row
   */
  public TypedExcelTableBean(List<Object> colList) {
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

        if (fieldName == null) {
          continue;
        }

        Field field = null;
        Class<?> clazz = this.getClass();
        while (clazz != null) {
          try {
            field = clazz.getDeclaredField(fieldName);
            break;
          } catch (NoSuchFieldException ignored) {
            clazz = clazz.getSuperclass();
            if (clazz == null) {
              throw new RuntimeException("Trying to set a value to the field in the bean, "
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

  private @Nullable Object convertToFieldType(Class<?> fieldType, @Nullable Object value,
      String fieldName) {
    if (value == null) {
      return null;
    }

    if (value instanceof Double d) {
      return convertDouble(fieldType, d, fieldName);
    }
    if (value instanceof LocalDate ld) {
      return convertLocalDate(fieldType, ld, fieldName);
    }
    if (value instanceof LocalDateTime ldt) {
      return convertLocalDateTime(fieldType, ldt, fieldName);
    }
    if (value instanceof String s) {
      return convertString(fieldType, s, fieldName);
    }
    if (value instanceof Boolean b) {
      return convertBoolean(fieldType, b, fieldName);
    }

    throw new RuntimeException("Unsupported cell value type: " + value.getClass().getSimpleName()
        + " for field '" + fieldName + "'.");
  }

  private @Nullable Object convertDouble(Class<?> fieldType, double d, String fieldName) {
    if (fieldType == String.class) {
      return formatDouble(d);
    }
    if (fieldType == Double.class || fieldType == double.class) {
      return d;
    }
    if (fieldType == Float.class || fieldType == float.class) {
      return (float) d;
    }
    if (fieldType == Integer.class || fieldType == int.class) {
      return (int) Math.round(d);
    }
    if (fieldType == Long.class || fieldType == long.class) {
      return Math.round(d);
    }
    if (fieldType == Short.class || fieldType == short.class) {
      return (short) Math.round(d);
    }
    if (fieldType == BigDecimal.class) {
      return BigDecimal.valueOf(d);
    }
    if (fieldType == BigInteger.class) {
      return BigInteger.valueOf(Math.round(d));
    }
    throw new RuntimeException("Cannot convert numeric (Double) value '" + d
        + "' to field type " + fieldType.getSimpleName() + " for field '" + fieldName + "'.");
  }

  private @Nullable Object convertLocalDate(Class<?> fieldType, LocalDate ld, String fieldName) {
    if (fieldType == LocalDate.class) {
      return ld;
    }
    if (fieldType == LocalDateTime.class) {
      return ld.atStartOfDay();
    }
    if (fieldType == String.class) {
      return ld.toString();
    }
    throw new RuntimeException("Cannot convert LocalDate value '" + ld
        + "' to field type " + fieldType.getSimpleName() + " for field '" + fieldName + "'.");
  }

  private @Nullable Object convertLocalDateTime(Class<?> fieldType, LocalDateTime ldt,
      String fieldName) {
    if (fieldType == LocalDateTime.class) {
      return ldt;
    }
    if (fieldType == LocalDate.class) {
      return ldt.toLocalDate();
    }
    if (fieldType == LocalTime.class) {
      return ldt.toLocalTime();
    }
    if (fieldType == String.class) {
      return ldt.toString();
    }
    throw new RuntimeException("Cannot convert LocalDateTime value '" + ldt
        + "' to field type " + fieldType.getSimpleName() + " for field '" + fieldName + "'.");
  }

  private @Nullable Object convertString(Class<?> fieldType, String s, String fieldName) {
    if (fieldType == String.class) {
      return s;
    }
    if (fieldType == Boolean.class || fieldType == boolean.class) {
      return Boolean.valueOf(s);
    }
    throw new RuntimeException("Cannot convert String value '" + s
        + "' to field type " + fieldType.getSimpleName() + " for field '" + fieldName + "'.");
  }

  private @Nullable Object convertBoolean(Class<?> fieldType, boolean b, String fieldName) {
    if (fieldType == Boolean.class || fieldType == boolean.class) {
      return b;
    }
    if (fieldType == String.class) {
      return Boolean.toString(b);
    }
    throw new RuntimeException("Cannot convert Boolean value '" + b
        + "' to field type " + fieldType.getSimpleName() + " for field '" + fieldName + "'.");
  }

  private static String formatDouble(double d) {
    long rounded = Math.round(d);
    if (d == (double) rounded) {
      return Long.toString(rounded);
    }
    return Double.toString(d);
  }
}
