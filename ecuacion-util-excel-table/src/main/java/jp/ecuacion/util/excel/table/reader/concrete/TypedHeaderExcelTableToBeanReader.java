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
package jp.ecuacion.util.excel.table.reader.concrete;

import jakarta.validation.Validation;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.Violations;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Reads an Excel table with one or multiple header rows and stores each data row into a bean.
 *
 * <p>Cell values are mapped to native Java types before being set on the bean fields:
 *     {@link Double} for numeric cells, {@link java.time.LocalDate} or
 *     {@link java.time.LocalDateTime} for date-formatted cells, {@link String} for string cells,
 *     and {@link Boolean} for boolean cells. The bean's field type drives any further
 *     coercion (e.g. {@link Double} → {@link Integer} via {@link Math#round(double)}).</p>
 *
 * @param <T> the bean type, must extend {@link TypedExcelTableBean}
 */
public class TypedHeaderExcelTableToBeanReader<T extends TypedExcelTableBean>
    extends TypedHeaderExcelTableReader {

  private Class<?> beanClass;

  /**
   * Stores the 1-based Excel row number where data starts (first row after the header).
   */
  protected int dataStartExcelRowNumber = 0;

  /**
   * Constructs a new instance with multiple header rows.
   *
   * <p>Defaults: {@code tableStartRowNumber = null} (auto-detect by header label),
   *     {@code tableStartColumnNumber = 1}, {@code tableRowSize = null}.</p>
   *
   * @param beanClass the class of the bean ({@code T}) — pass explicitly because Java generics
   *     do not allow {@code T.class}
   * @param sheetName sheet name
   * @param headerLabels expected header labels: {@code headerLabels[row][col]}, top row first
   */
  public TypedHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[][] headerLabels) {
    super(sheetName, headerLabels);
    this.beanClass = beanClass;
  }

  /**
   * Reads the Excel table, converts each row to a bean, and validates.
   *
   * @param filePath path to the Excel file
   * @return list of beans
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public List<T> readToBean(String filePath) throws EncryptedDocumentException, IOException {
    return readToBean(filePath, true);
  }

  /**
   * Reads the Excel table and converts each row to a bean.
   *
   * @param filePath path to the Excel file
   * @param validates whether to run Jakarta Validation on each bean
   * @return list of beans
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public List<T> readToBean(String filePath, boolean validates)
      throws EncryptedDocumentException, IOException {
    final String msgId = "jp.ecuacion.util.excel.reader.ValidationMessagePostfix.message";
    List<T> rtnList = excelTableToBeanList(filePath);

    if (validates) {
      for (int i = 0; i < rtnList.size(); i++) {
        T bean = rtnList.get(i);
        int excelRowNumber = dataStartExcelRowNumber + i;
        new Violations()
            .addAll(Validation.buildDefaultValidatorFactory().getValidator().validate(bean))
            .messageParameters(Violations.newMessageParameters().isMessageWithItemName(true)
                .messagePostfix(Arg.message(msgId, getSheetName(), String.valueOf(excelRowNumber))))
            .throwIfAny();

        bean.afterReading();
      }
    }

    return rtnList;
  }

  /**
   * Reads the Excel file and converts rows to beans.
   *
   * <p>Override to supply a fixed list for testing validation logic without Excel files.
   *     When overriding, set {@link #dataStartExcelRowNumber} explicitly if row numbers
   *     are needed in validation messages.</p>
   *
   * @param filePath path to the Excel file
   * @return list of beans
   * @throws IOException IOException
   */
  protected List<T> excelTableToBeanList(String filePath) throws IOException {
    try (Workbook workbook = ExcelReadUtil.openForRead(filePath)) {
      List<List<Object>> lines = read(workbook);

      Sheet sheet = workbook.getSheet(getSheetName());
      int poiBasisHeaderRow =
          getPoiBasisDeterminedTableStartRowNumber(sheet, tableStartColumnNumber);
      dataStartExcelRowNumber = poiBasisHeaderRow + getNumberOfHeaderLines() + 1;

      boolean usesAnnotation = usesExcelColumnAnnotation(beanClass);
      List<T> rtnList = new ArrayList<>();
      for (List<Object> line : lines) {
        try {
          List<Object> colList = usesAnnotation ? buildReorderedColList(line) : line;
          @SuppressWarnings("unchecked")
          T bean = (T) beanClass.getConstructor(List.class).newInstance(colList);
          rtnList.add(bean);
        } catch (Exception ex) {
          throw new RuntimeException(ex);
        }
      }
      return rtnList;
    }
  }

  private boolean usesExcelColumnAnnotation(Class<?> clazz) {
    Class<?> current = clazz;
    while (current != null && current != TypedExcelTableBean.class) {
      for (Field f : current.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          return true;
        }
      }
      current = current.getSuperclass();
    }
    return false;
  }

  private List<Object> buildReorderedColList(List<Object> colList) {
    String[][] h = getHeaderLabelData();
    int numCols = getHeaderLabels().length;

    List<Class<?>> hierarchy = buildClassHierarchy(beanClass);

    List<Object> reordered = new ArrayList<>();
    for (Class<?> c : hierarchy) {
      for (Field f : c.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          @SuppressWarnings("null")
          String[] annotLabels = f.getAnnotation(ExcelColumn.class).value();
          int colIdx = findColumnIndex(h, numCols, annotLabels);
          if (colIdx < 0) {
            throw new RuntimeException("@ExcelColumn " + Arrays.toString(annotLabels)
                + " not found in headerLabels of " + getSheetName() + ".");
          }
          reordered.add(colList.get(colIdx));
        }
      }
    }
    return reordered;
  }

  private int findColumnIndex(String[][] headerLabels2d, int numCols, String[] annotLabels) {
    for (int i = 0; i < numCols; i++) {
      if (columnMatches(headerLabels2d, i, annotLabels)) {
        return i;
      }
    }
    return -1;
  }

  private boolean columnMatches(String[][] headerLabels2d, int colIdx, String[] annotLabels) {
    if (annotLabels.length == 1) {
      for (String[] headerRow : headerLabels2d) {
        if (!annotLabels[0].equals(headerRow[colIdx])) {
          return false;
        }
      }
      return true;
    }
    if (annotLabels.length != headerLabels2d.length) {
      return false;
    }
    for (int rowIdx = 0; rowIdx < headerLabels2d.length; rowIdx++) {
      if (!annotLabels[rowIdx].equals(headerLabels2d[rowIdx][colIdx])) {
        return false;
      }
    }
    return true;
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
  public TypedHeaderExcelTableToBeanReader<T> tableStartRowNumber(@Nullable Integer value) {
    return (TypedHeaderExcelTableToBeanReader<T>) super.tableStartRowNumber(value);
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableToBeanReader<T> tableStartColumnNumber(int value) {
    return (TypedHeaderExcelTableToBeanReader<T>) super.tableStartColumnNumber(value);
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableToBeanReader<T> tableRowSize(@Nullable Integer value) {
    return (TypedHeaderExcelTableToBeanReader<T>) super.tableRowSize(value);
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableToBeanReader<T> tableColumnSize(@Nullable Integer value) {
    return (TypedHeaderExcelTableToBeanReader<T>) super.tableColumnSize(value);
  }

  @Override
  public TypedHeaderExcelTableToBeanReader<T> withIgnoresAdditionalColumnsOfHeaderData(
      boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @SuppressWarnings("unchecked")
  @Override
  public TypedHeaderExcelTableToBeanReader<T> withVerticalAndHorizontalOpposite(boolean value) {
    return (TypedHeaderExcelTableToBeanReader<T>) super.withVerticalAndHorizontalOpposite(value);
  }
}
