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
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;
import org.apache.poi.EncryptedDocumentException;
import org.jspecify.annotations.Nullable;

/**
 * Writes tables with two or more header rows from a list of {@link TypedExcelTableBean} instances.
 *
 * <p>Each bean field's value is written to Excel using its native Java type — for example a
 *     {@link java.time.LocalDate} field is written as a date-formatted cell, not as a string.
 *     See {@link jp.ecuacion.util.excel.table.writer.IfDataTypeTypedExcelTableWriter} for
 *     details on how each value type is written.</p>
 *
 * <p>Column-to-field mapping uses {@link ExcelColumn} annotations when present,
 *     or {@link TypedExcelTableBean#getFieldNameArray()} for positional mapping.</p>
 *
 * <p>For single-row headers use {@link TypedOneLineHeaderExcelTableFromBeanWriter}.</p>
 *
 * @param <T> the bean type, must extend {@link TypedExcelTableBean}
 */
public class TypedHeaderExcelTableFromBeanWriter<T extends TypedExcelTableBean>
    extends TypedHeaderExcelTableWriter {

  /**
   * Constructs a new instance with the sheet name and multiple header rows.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param headerLabels expected labels for each header row: {@code headerLabels[row][col]},
   *     top row first
   */
  public TypedHeaderExcelTableFromBeanWriter(String sheetName, String[][] headerLabels) {
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
    List<List<Object>> data = new ArrayList<>();
    for (T bean : beans) {
      data.add(beanToObjectList(bean));
    }
    write(templateFilePath, destFilePath, data);
  }

  private List<Object> beanToObjectList(T bean) {
    List<Class<?>> hierarchy = buildClassHierarchy(bean.getClass());
    boolean usesAnnotation = usesExcelColumnAnnotation(hierarchy);

    try {
      if (usesAnnotation) {
        return buildObjectListByAnnotation(bean, hierarchy);
      } else {
        return buildObjectListByFieldOrder(bean, hierarchy);
      }
    } catch (RuntimeException ex) {
      throw ex;
    } catch (Exception ex) {
      throw new RuntimeException(ex);
    }
  }

  private List<Object> buildObjectListByAnnotation(T bean, List<Class<?>> hierarchy)
      throws Exception {
    String[][] headerData = getHeaderLabelData();
    int numCols = getHeaderLabels().length;
    List<Object> result = new ArrayList<>();
    for (int colIdx = 0; colIdx < numCols; colIdx++) {
      Field field = findFieldForColumn(hierarchy, headerData, colIdx);
      if (field != null) {
        field.setAccessible(true);
        result.add(field.get(bean));
      } else {
        result.add(null);
      }
    }
    return result;
  }

  private List<Object> buildObjectListByFieldOrder(T bean, List<Class<?>> hierarchy)
      throws Exception {
    Method method = TypedExcelTableBean.class.getDeclaredMethod("getFieldNameArray");
    method.setAccessible(true);
    String[] fieldNames = (String[]) method.invoke(bean);

    List<Object> result = new ArrayList<>();
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
        result.add(field.get(bean));
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
    while (clazz != null && clazz != TypedExcelTableBean.class) {
      hierarchy.add(0, clazz);
      clazz = clazz.getSuperclass();
    }
    return hierarchy;
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableFromBeanWriter<T> defaultDateFormat(String formatPattern) {
    return (TypedHeaderExcelTableFromBeanWriter<T>) super.defaultDateFormat(formatPattern);
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableFromBeanWriter<T> defaultDateTimeFormat(String formatPattern) {
    return (TypedHeaderExcelTableFromBeanWriter<T>) super.defaultDateTimeFormat(formatPattern);
  }
}
