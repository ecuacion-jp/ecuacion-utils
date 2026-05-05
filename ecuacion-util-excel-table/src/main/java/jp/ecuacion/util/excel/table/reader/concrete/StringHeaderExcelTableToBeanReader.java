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

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.Violations;
import jp.ecuacion.util.excel.enums.NoDataString;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.StringExcelTableBean;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jspecify.annotations.Nullable;

/**
 * Reads an Excel table with one or multiple header rows and stores each data row into a bean.
 *
 * <p>This class unifies the former {@code StringOneLineHeaderExcelTableToBeanReader}
 *     and extends it with multi-row header support.
 *     Constructors accepting {@code String[]} provide backward compatibility
 *     for single-row headers; constructors accepting {@code String[][]} enable multi-row.</p>
 *
 * @param <T> the bean type, must extend {@link StringExcelTableBean}
 */
public class StringHeaderExcelTableToBeanReader<T extends StringExcelTableBean>
    extends StringHeaderExcelTableReader {

  private Class<?> beanClass;

  /**
   * Stores the 1-based Excel row number where data starts (first row after the header).
   *
   * <p>Set during {@link #excelTableToBeanList(String)} and used in {@link #readToBean(String,
   * boolean)} to include the Excel row number in validation error messages.</p>
   */
  protected int dataStartExcelRowNumber = 0;

  // ── single-row constructors ────────────────────────────────────────────────

  /**
   * Constructs a new instance with a single header row.
   * The obtained value from an empty cell is {@code null}.
   *
   * @param beanClass the class of the bean ({@code T}) — pass explicitly because Java generics
   *     do not allow {@code T.class}
   * @param sheetName sheet name
   * @param headerLabels expected header labels
   * @param tableStartRowNumber 1-based row number of the header, or {@code null} for
   *     auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum data rows, or {@code null} for auto-detection
   * @param parameterClass dummy varargs used only for type inference; leave empty
   */
  public StringHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[] headerLabels, @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, @SuppressWarnings("unchecked") T... parameterClass) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize);
    this.beanClass = beanClass;
  }

  /**
   * Constructs a new instance with a single header row and a specified empty-cell value.
   *
   * @param beanClass the class of the bean ({@code T})
   * @param sheetName sheet name
   * @param headerLabels expected header labels
   * @param tableStartRowNumber 1-based row number of the header, or {@code null} for
   *     auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum data rows, or {@code null} for auto-detection
   * @param noDataString the value returned for an empty cell
   */
  public StringHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[] headerLabels, @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, NoDataString noDataString) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        noDataString);
    this.beanClass = beanClass;
  }

  // ── multi-row constructors ─────────────────────────────────────────────────

  /**
   * Constructs a new instance with multiple header rows.
   * The obtained value from an empty cell is {@code null}.
   *
   * @param beanClass the class of the bean ({@code T})
   * @param sheetName sheet name
   * @param headerLabels expected header labels: {@code headerLabels[row][col]}, top row first
   * @param tableStartRowNumber 1-based row number of the first header row,
   *     or {@code null} for auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum data rows, or {@code null} for auto-detection
   */
  public StringHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[][] headerLabels, @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize);
    this.beanClass = beanClass;
  }

  /**
   * Constructs a new instance with multiple header rows and a specified empty-cell value.
   *
   * @param beanClass the class of the bean ({@code T})
   * @param sheetName sheet name
   * @param headerLabels expected header labels: {@code headerLabels[row][col]}, top row first
   * @param tableStartRowNumber 1-based row number of the first header row,
   *     or {@code null} for auto-detection
   * @param tableStartColumnNumber 1-based column number where the table starts
   * @param tableRowSize maximum data rows, or {@code null} for auto-detection
   * @param noDataString the value returned for an empty cell
   */
  public StringHeaderExcelTableToBeanReader(Class<?> beanClass, String sheetName,
      String[][] headerLabels, @Nullable Integer tableStartRowNumber, int tableStartColumnNumber,
      @Nullable Integer tableRowSize, NoDataString noDataString) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        noDataString);
    this.beanClass = beanClass;
  }

  // ── read methods ───────────────────────────────────────────────────────────

  /**
   * Reads the Excel table, converts each row to a bean, and validates.
   *
   * @param filePath path to the Excel file
   * @return list of beans
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public List<T> readToBean(String filePath)
      throws EncryptedDocumentException, IOException {
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
            .messageParameters(Violations.newMessageParameters()
                .isMessageWithItemName(true)
                .messagePostfix(Arg.message(msgId,
                    Arg.strings(getSheetName(), String.valueOf(excelRowNumber)))))
            .throwIfAny();

        bean.afterReading();
      }
    }

    return rtnList;
  }

  /**
   * Reads the Excel file and converts rows to beans.
   *
   * <p>Override this method to supply a fixed list for testing validation logic
   *     without preparing actual Excel files.
   *     When overriding, set {@link #dataStartExcelRowNumber} explicitly if row numbers
   *     are needed in validation messages.</p>
   *
   * @param filePath path to the Excel file
   * @return list of beans
   * @throws IOException IOException
   */
  protected List<T> excelTableToBeanList(String filePath) throws IOException {
    try (Workbook workbook = ExcelReadUtil.openForRead(filePath)) {
      List<List<String>> lines = read(workbook);

      Sheet sheet = workbook.getSheet(getSheetName());
      int poiBasisHeaderRow =
          getPoiBasisDeterminedTableStartRowNumber(sheet, tableStartColumnNumber);
      dataStartExcelRowNumber = poiBasisHeaderRow + getNumberOfHeaderLines() + 1;

      boolean usesAnnotation = usesExcelColumnAnnotation(beanClass);
      List<T> rtnList = new ArrayList<>();
      for (List<String> line : lines) {
        try {
          List<String> colList = usesAnnotation ? buildReorderedColList(line) : line;
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

  // ── @ExcelColumn matching ──────────────────────────────────────────────────

  private boolean usesExcelColumnAnnotation(Class<?> clazz) {
    Class<?> current = clazz;
    while (current != null && current != StringExcelTableBean.class) {
      for (Field f : current.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          return true;
        }
      }
      current = current.getSuperclass();
    }
    return false;
  }

  /**
   * Builds a column-value list ordered by {@link ExcelColumn} field declaration,
   *     matched to the header labels by annotation value.
   *
   * <p>For multi-row headers, the annotation value array is matched against
   *     the corresponding header row values for each column.
   *     A single-element annotation matches any column where all header rows
   *     have that same value (vertically merged).</p>
   *
   * @param colList column values in table-column order
   * @return reordered list aligned to the {@link ExcelColumn} field scan order
   */
  private List<String> buildReorderedColList(List<String> colList) {
    String[][] h = getHeaderLabels2d();
    int numCols = getHeaderLabels().length;

    List<Class<?>> hierarchy = buildClassHierarchy(beanClass);

    List<String> reordered = new ArrayList<>();
    for (Class<?> c : hierarchy) {
      for (Field f : c.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          String[] annotLabels = f.getAnnotation(ExcelColumn.class).value();
          int colIdx = findColumnIndex(h, numCols, annotLabels);
          if (colIdx < 0) {
            throw new RuntimeException("@ExcelColumn " + java.util.Arrays.toString(annotLabels)
                + " not found in headerLabels of " + getSheetName() + ".");
          }
          reordered.add(colList.get(colIdx));
        }
      }
    }
    return reordered;
  }

  /**
   * Finds the 0-based column index whose header-label key matches {@code annotLabels}.
   *
   * @param headerLabels2d all header rows
   * @param numCols number of columns
   * @param annotLabels {@link ExcelColumn#value()}
   * @return 0-based column index, or {@code -1} if not found
   */
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
      // Single label: match if all header rows have the same value.
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
    while (clazz != null && clazz != StringExcelTableBean.class) {
      hierarchy.add(0, clazz);
      clazz = clazz.getSuperclass();
    }
    return hierarchy;
  }

  // ── highlightErrors ────────────────────────────────────────────────────────

  /**
   * Writes a copy of the original Excel file with error cells highlighted in red.
   *
   * <p>Typical usage:</p>
   * <pre>{@code
   * try {
   *   reader.readToBean(filePath);
   * } catch (ViolationException ex) {
   *   reader.highlightErrors(filePath, ex.getViolations(), outputPath);
   * }
   * }</pre>
   *
   * <p>When the bean uses {@link ExcelColumn} annotations, only the violated cells
   *     are highlighted; otherwise, all data cells in the violated row are highlighted.</p>
   *
   * @param originalPath path to the source Excel file
   * @param violations violations from a caught
   *     {@link jp.ecuacion.lib.core.exception.ViolationException}
   * @param outputPath path where the highlighted Excel file is saved
   * @throws IOException IOException
   * @throws IllegalArgumentException if {@code violations} does not contain cell location info
   */
  public void highlightErrors(String originalPath, Violations violations, String outputPath)
      throws IOException {
    Arg postfix = violations.messageParameters().getMessagePostfix();
    if (postfix == null || postfix.getMessageArgs().length < 2) {
      throw new IllegalArgumentException(
          "Violations does not contain cell location info. "
              + "Make sure the violations are from readToBean().");
    }

    int excelRowNumber = Integer.parseInt(postfix.getMessageArgs()[1].getArgString());

    Set<String> violatedFieldNames = new LinkedHashSet<>();
    for (ConstraintViolation<?> cv : violations.getConstraintViolations()) {
      String path = cv.getPropertyPath().toString();
      int dotIdx = path.lastIndexOf('.');
      violatedFieldNames.add(dotIdx >= 0 ? path.substring(dotIdx + 1) : path);
    }

    try (Workbook workbook = WorkbookFactory.create(new File(originalPath));
        FileOutputStream fos = new FileOutputStream(outputPath)) {
      Sheet sheet = workbook.getSheet(getSheetName());
      if (sheet == null) {
        throw new RuntimeException("Sheet not found: " + getSheetName());
      }
      int poiRowIndex = excelRowNumber - 1;
      Row row = sheet.getRow(poiRowIndex);
      if (row == null) {
        row = sheet.createRow(poiRowIndex);
      }

      CellStyle errorStyle = workbook.createCellStyle();
      errorStyle.setFillForegroundColor(IndexedColors.RED1.getIndex());
      errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

      if (usesExcelColumnAnnotation(beanClass)) {
        for (int poiColIdx : resolvePoiColumnIndices(violatedFieldNames)) {
          Cell cell = row.getCell(poiColIdx);
          if (cell == null) {
            cell = row.createCell(poiColIdx);
          }
          cell.setCellStyle(errorStyle);
        }
      } else {
        for (int i = 0; i < getHeaderLabels().length; i++) {
          int poiColIdx = tableStartColumnNumber - 1 + i;
          Cell cell = row.getCell(poiColIdx);
          if (cell == null) {
            cell = row.createCell(poiColIdx);
          }
          cell.setCellStyle(errorStyle);
        }
      }

      workbook.write(fos);
    }
  }

  private List<Integer> resolvePoiColumnIndices(Set<String> fieldNames) {
    String[][] h = getHeaderLabels2d();
    int numCols = getHeaderLabels().length;

    List<Class<?>> hierarchy = buildClassHierarchy(beanClass);
    java.util.Map<String, Integer> fieldToColIdx = new java.util.HashMap<>();
    for (Class<?> c : hierarchy) {
      for (Field f : c.getDeclaredFields()) {
        if (f.isAnnotationPresent(ExcelColumn.class)) {
          String[] annotLabels = f.getAnnotation(ExcelColumn.class).value();
          int colIdx = findColumnIndex(h, numCols, annotLabels);
          if (colIdx >= 0) {
            fieldToColIdx.put(f.getName(), tableStartColumnNumber - 1 + colIdx);
          }
        }
      }
    }

    List<Integer> result = new ArrayList<>();
    for (String fieldName : fieldNames) {
      Integer poiColIdx = fieldToColIdx.get(fieldName);
      if (poiColIdx != null) {
        result.add(poiColIdx);
      }
    }
    return result;
  }

  // ── method chaining overrides ──────────────────────────────────────────────

  @SuppressWarnings("unchecked")
  @Override
  public StringHeaderExcelTableToBeanReader<T> defaultDateTimeFormat(
      DateTimeFormatter dateTimeFormat) {
    return (StringHeaderExcelTableToBeanReader<T>) super.defaultDateTimeFormat(dateTimeFormat);
  }

  @SuppressWarnings("unchecked")
  @Override
  public StringHeaderExcelTableToBeanReader<T> columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringHeaderExcelTableToBeanReader<T>) super.columnDateTimeFormat(columnNumber,
        dateTimeFormat);
  }

  @Override
  public StringHeaderExcelTableToBeanReader<T> ignoresAdditionalColumnsOfHeaderData(
      boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }
}
