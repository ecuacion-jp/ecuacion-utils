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

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("TypedOneLineHeaderExcelTableToBeanReader / TypedHeaderExcelTableToBeanReader")
public class TypedHeaderExcelTableToBeanReaderTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  // --- test beans ---

  /** Bean with native-typed fields, mapped by field order. */
  static class PersonBean extends TypedExcelTableBean {
    @NotBlank @Nullable String name;
    @Nullable Integer age;
    @Nullable Double score;
    @Nullable LocalDate birthDate;
    @Nullable Boolean active;

    public PersonBean(List<Object> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"name", "age", "score", "birthDate", "active"};
    }
  }

  /** Bean using {@code @ExcelColumn}. */
  static class AnnotatedBean extends TypedExcelTableBean {
    @ExcelColumn("name") @NotBlank @Nullable String name;
    @ExcelColumn("age") @Min(1) @Nullable Integer age;

    public AnnotatedBean(List<Object> colList) {
      super(colList);
    }
  }

  /** Bean whose field is a date-time field, used to check LocalDateTime mapping. */
  static class EventBean extends TypedExcelTableBean {
    @Nullable String title;
    @Nullable LocalDateTime startsAt;

    public EventBean(List<Object> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"title", "startsAt"};
    }
  }

  // --- helpers ---

  private static Row getOrCreateRow(Sheet sheet, int poiRow) {
    Row row = sheet.getRow(poiRow);
    return row == null ? sheet.createRow(poiRow) : row;
  }

  private static void setStringCell(Sheet sheet, int poiRow, int poiCol, String value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setNumericCell(Sheet sheet, int poiRow, int poiCol, double value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setBooleanCell(Sheet sheet, int poiRow, int poiCol, boolean value) {
    getOrCreateRow(sheet, poiRow).createCell(poiCol).setCellValue(value);
  }

  private static void setDateFormattedCell(Workbook wb, Sheet sheet, int poiRow, int poiCol,
      LocalDateTime value, String formatPattern) {
    Cell cell = getOrCreateRow(sheet, poiRow).createCell(poiCol);
    CellStyle style = wb.createCellStyle();
    style.setDataFormat(wb.createDataFormat().getFormat(formatPattern));
    cell.setCellStyle(style);
    cell.setCellValue(value);
  }

  private Path writeTempExcel(Workbook wb) throws Exception {
    Path file = tempDir.resolve("test.xlsx");
    try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
      wb.write(fos);
    }
    return file;
  }

  @Nested
  @DisplayName("normal case: stored into the Bean with native types")
  class Normal {

    @Test
    @DisplayName("String / Double / LocalDate / Boolean fields are each stored with their own type")
    void nativeTypesStoredAsIs() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 0, 2, "score");
        setStringCell(sheet, 0, 3, "birthDate");
        setStringCell(sheet, 0, 4, "active");

        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, 25.0);
        setNumericCell(sheet, 1, 2, 92.5);
        setDateFormattedCell(wb, sheet, 1, 3, LocalDateTime.of(2000, 4, 1, 0, 0), "yyyy-mm-dd");
        setBooleanCell(sheet, 1, 4, true);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<PersonBean>(PersonBean.class,
            "Sheet1", new String[] {"name", "age", "score", "birthDate", "active"})
            .tableStartRowNumber(1);
        List<PersonBean> result = reader.readToBean(file.toString(), false);

        assertThat(result).hasSize(1);
        PersonBean bean = result.get(0);
        assertThat(bean.name).isEqualTo("Alice");
        assertThat(bean.age).isEqualTo(25);
        assertThat(bean.score).isEqualTo(92.5);
        assertThat(bean.birthDate).isEqualTo(LocalDate.of(2000, 4, 1));
        assertThat(bean.active).isTrue();
      }
    }

    @Test
    @DisplayName("LocalDateTime field is stored as-is from a date cell with a time component")
    void localDateTimeField() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "title");
        setStringCell(sheet, 0, 1, "startsAt");
        setStringCell(sheet, 1, 0, "Conference");
        setDateFormattedCell(wb, sheet, 1, 1, LocalDateTime.of(2026, 1, 15, 9, 30),
            "yyyy-mm-dd hh:mm:ss");
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<EventBean>(EventBean.class,
            "Sheet1", new String[] {"title", "startsAt"}).tableStartRowNumber(1);
        List<EventBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).title).isEqualTo("Conference");
        assertThat(result.get(0).startsAt).isEqualTo(LocalDateTime.of(2026, 1, 15, 9, 30));
      }
    }
  }

  @Nested
  @DisplayName("numeric-to-integer field rounding conversion")
  class IntegerRounding {

    @Test
    @DisplayName("whole number value in Excel (actually 1.0) → stored as 1 in an Integer field")
    void wholeNumberDouble() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 0, 2, "score");
        setStringCell(sheet, 0, 3, "birthDate");
        setStringCell(sheet, 0, 4, "active");

        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, 1.0); // Excel displays "1", actual double is 1.0
        setNumericCell(sheet, 1, 2, 1.0);
        setDateFormattedCell(wb, sheet, 1, 3, LocalDateTime.of(2000, 4, 1, 0, 0), "yyyy-mm-dd");
        setBooleanCell(sheet, 1, 4, true);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<PersonBean>(PersonBean.class,
            "Sheet1", new String[] {"name", "age", "score", "birthDate", "active"})
            .tableStartRowNumber(1);
        List<PersonBean> result = reader.readToBean(file.toString(), false);

        // Integer field gets rounded to 1 (not the buggy "1.0" string parse failure
        // that the String-based reader would produce).
        assertThat(result.get(0).age).isEqualTo(1);
        // Double field keeps the value as-is.
        assertThat(result.get(0).score).isEqualTo(1.0);
      }
    }

    @Test
    @DisplayName("fractional value → stored in an Integer field rounded via Math.round")
    void roundsFractionalValue() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 0, 2, "score");
        setStringCell(sheet, 0, 3, "birthDate");
        setStringCell(sheet, 0, 4, "active");

        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, 1.6); // rounds up to 2
        setNumericCell(sheet, 1, 2, 1.6);
        setDateFormattedCell(wb, sheet, 1, 3, LocalDateTime.of(2000, 4, 1, 0, 0), "yyyy-mm-dd");
        setBooleanCell(sheet, 1, 4, true);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<PersonBean>(PersonBean.class,
            "Sheet1", new String[] {"name", "age", "score", "birthDate", "active"})
            .tableStartRowNumber(1);
        List<PersonBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).age).isEqualTo(Math.round(1.6));
        assertThat(result.get(0).score).isEqualTo(1.6);
      }
    }
  }

  @Nested
  @DisplayName("@ExcelColumn annotation")
  class ExcelColumnAnnotation {

    @Test
    @DisplayName("even when the Excel column order is reversed from the header, "
        + "@ExcelColumn maps by matching label")
    void columnOrderIndependent() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "age");
        setStringCell(sheet, 0, 1, "name");
        setNumericCell(sheet, 1, 0, 25.0);
        setStringCell(sheet, 1, 1, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(AnnotatedBean.class,
            "Sheet1", new String[] {"age", "name"}).tableStartRowNumber(1);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }
  }

  @Nested
  @DisplayName("validation")
  class Validation {

    @Test
    @DisplayName("@Min(1) violation → throws ViolationException")
    void violationThrowsException() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, -1.0); // @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(AnnotatedBean.class,
            "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);

        assertThatThrownBy(() -> reader.readToBean(file.toString(), true))
            .isInstanceOf(ViolationException.class);
      }
    }

    @Test
    @DisplayName("validates=false → no exception is thrown even if there is a violation")
    void validatesFalseSkipsValidation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 1, 0, "Alice");
        setNumericCell(sheet, 1, 1, -1.0);
        Path file = writeTempExcel(wb);

        var reader = new TypedOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(AnnotatedBean.class,
            "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).age).isEqualTo(-1);
      }
    }
  }
}
