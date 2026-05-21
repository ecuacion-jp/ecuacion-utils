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
package jp.ecuacion.util.excel.table.writer.concrete;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.StringExcelTableBean;
import org.apache.poi.EncryptedDocumentException;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with two or more header rows from a list of {@link StringExcelTableBean} instances.
 *
 * <p>Column-to-field mapping uses {@link ExcelColumn} annotations when present,
 *     or {@link StringExcelTableBean#getFieldNameArray()} for positional mapping.</p>
 *
 * <p>For single-row headers use {@link StringOneLineHeaderExcelTableFromBeanWriter}.</p>
 *
 * @param <T> the bean type, must extend {@link StringExcelTableBean}
 */
public class StringHeaderExcelTableFromBeanWriter<T extends StringExcelTableBean>
    extends StringHeaderExcelTableWriter {

  private DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ISO_LOCAL_DATE;

  /**
   * Constructs a new instance with the sheet name and multiple header rows.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected labels for each header row: {@code headerLabels[row][col]},
   *     top row first
   */
  public StringHeaderExcelTableFromBeanWriter(String sheetName, String[][] headerLabels) {
    super(sheetName, headerLabels);
  }

  /**
   * Writes data from a list of beans into the template file.
   *
   * @param templateFilePath path to the template Excel file
   * @param destFilePath path to the output file
   * @param beans list of beans to write
   * @throws ExcelTableException if a header mismatch is detected
   * @throws EncryptedDocumentException if the file is encrypted
   * @throws IOException if an I/O error occurs
   */
  public void writeFromBean(String templateFilePath, String destFilePath, List<T> beans)
      throws ExcelTableException, EncryptedDocumentException, IOException {
    List<List<String>> data = new ArrayList<>();
    for (T bean : beans) {
      data.add(beanToStringList(bean));
    }
    write(templateFilePath, destFilePath, data);
  }

  /**
   * Sets the {@link DateTimeFormatter} used to convert date/time field values to strings.
   * Defaults to {@link DateTimeFormatter#ISO_LOCAL_DATE} ({@code yyyy-MM-dd}).
   *
   * @param formatter the formatter to use
   * @return this writer
   */
  public StringHeaderExcelTableFromBeanWriter<T> defaultDateTimeFormat(
      DateTimeFormatter formatter) {
    this.dateTimeFormatter = formatter;
    return this;
  }

  private List<String> beanToStringList(T bean) {
    List<Class<?>> hierarchy = buildClassHierarchy(bean.getClass());
    boolean usesAnnotation = usesExcelColumnAnnotation(hierarchy);

    try {
      if (usesAnnotation) {
        return buildStringListByAnnotation(bean, hierarchy);
      } else {
        return buildStringListByFieldOrder(bean, hierarchy);
      }
    } catch (RuntimeException ex) {
      throw ex;
    } catch (Exception ex) {
      throw new RuntimeException(ex);
    }
  }

  private List<String> buildStringListByAnnotation(T bean, List<Class<?>> hierarchy)
      throws Exception {
    String[][] headerData = getHeaderLabelData();
    int numCols = getHeaderLabels().length;
    List<String> result = new ArrayList<>();
    for (int colIdx = 0; colIdx < numCols; colIdx++) {
      Field field = findFieldForColumn(hierarchy, headerData, colIdx);
      if (field != null) {
        field.setAccessible(true);
        result.add(convertToString(field.get(bean)));
      } else {
        result.add(null);
      }
    }
    return result;
  }

  private List<String> buildStringListByFieldOrder(T bean, List<Class<?>> hierarchy)
      throws Exception {
    Method method = StringExcelTableBean.class.getDeclaredMethod("getFieldNameArray");
    method.setAccessible(true);
    String[] fieldNames = (String[]) method.invoke(bean);

    List<String> result = new ArrayList<>();
    for (String fieldName : Objects.requireNonNull(fieldNames)) {
      if (fieldName == null) {
        result.add(null);
      } else {
        Field field = findFieldByName(hierarchy, fieldName);
        if (field == null) {
          throw new RuntimeException("Field '" + fieldName + "' not found in bean class "
              + bean.getClass().getSimpleName() + ".");
        }
        field.setAccessible(true);
        result.add(convertToString(field.get(bean)));
      }
    }
    return result;
  }

  @Nullable
  private Field findFieldForColumn(List<Class<?>> hierarchy, String[][] headerData, int colIdx) {
    for (Class<?> clazz : hierarchy) {
      for (Field field : clazz.getDeclaredFields()) {
        if (field.isAnnotationPresent(ExcelColumn.class)) {
          String[] annotLabels =
              Objects.requireNonNull(field.getAnnotation(ExcelColumn.class)).value();
          if (columnMatchesAnnotation(headerData, colIdx, annotLabels)) {
            return field;
          }
        }
      }
    }
    return null;
  }

  private boolean columnMatchesAnnotation(String[][] headerData, int colIdx, String[] annotLabels) {
    if (annotLabels.length == 1) {
      for (String[] headerRow : headerData) {
        if (colIdx >= headerRow.length || !annotLabels[0].equals(headerRow[colIdx])) {
          return false;
        }
      }
      return true;
    }
    if (annotLabels.length != headerData.length) {
      return false;
    }
    for (int rowIdx = 0; rowIdx < headerData.length; rowIdx++) {
      if (colIdx >= headerData[rowIdx].length
          || !annotLabels[rowIdx].equals(headerData[rowIdx][colIdx])) {
        return false;
      }
    }
    return true;
  }

  @Nullable
  private Field findFieldByName(List<Class<?>> hierarchy, String fieldName) {
    for (Class<?> clazz : hierarchy) {
      try {
        return clazz.getDeclaredField(fieldName);
      } catch (NoSuchFieldException ignored) {
        // continue up hierarchy
      }
    }
    return null;
  }

  @Nullable
  private String convertToString(@Nullable Object value) {
    if (value == null) {
      return null;
    }
    if (value instanceof LocalDate date) {
      return date.format(dateTimeFormatter);
    }
    if (value instanceof LocalDateTime dateTime) {
      return dateTime.format(dateTimeFormatter);
    }
    if (value instanceof LocalTime time) {
      return time.toString();
    }
    return value.toString();
  }

  private boolean usesExcelColumnAnnotation(List<Class<?>> hierarchy) {
    for (Class<?> clazz : hierarchy) {
      for (Field field : clazz.getDeclaredFields()) {
        if (field.isAnnotationPresent(ExcelColumn.class)) {
          return true;
        }
      }
    }
    return false;
  }

  private List<Class<?>> buildClassHierarchy(Class<?> leaf) {
    List<Class<?>> hierarchy = new ArrayList<>();
    Class<?> clazz = leaf;
    while (clazz != null && clazz != StringExcelTableBean.class) {
      hierarchy.add(0, clazz);
      clazz = clazz.getSuperclass();
    }
    return hierarchy;
  }
}
