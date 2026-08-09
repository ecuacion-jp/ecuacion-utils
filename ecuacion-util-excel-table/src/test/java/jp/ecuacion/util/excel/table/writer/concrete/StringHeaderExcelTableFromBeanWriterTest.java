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

import static org.assertj.core.api.Assertions.assertThat;
import java.io.FileOutputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.StringExcelTableBean;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("StringOneLineHeaderExcelTableFromBeanWriter / StringHeaderExcelTableFromBeanWriter")
public class StringHeaderExcelTableFromBeanWriterTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  // --- test beans ---

  static class FieldOrderBean extends StringExcelTableBean {
    @Nullable String name;
    @Nullable Integer age;

    FieldOrderBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"name", "age"};
    }
  }

  static class AnnotatedBean extends StringExcelTableBean {
    @ExcelColumn("name") @Nullable String name;
    @ExcelColumn("age") @Nullable Integer age;

    AnnotatedBean(List<String> colList) {
      super(colList);
    }
  }

  static class AnnotatedBaseBean extends StringExcelTableBean {
    @ExcelColumn("id") @Nullable Integer id;

    AnnotatedBaseBean(List<String> colList) {
      super(colList);
    }
  }

  static class AnnotatedSubBean extends AnnotatedBaseBean {
    @ExcelColumn("name") @Nullable String name;

    AnnotatedSubBean(List<String> colList) {
      super(colList);
    }
  }

  static class DateBean extends StringExcelTableBean {
    @Nullable LocalDate date;

    DateBean(LocalDate date) {
      super(List.of(date.toString()));
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"date"};
    }
  }

  static class DateTimeBean extends StringExcelTableBean {
    @Nullable LocalDateTime dateTime;
    @Nullable LocalTime time;

    DateTimeBean(LocalDateTime dateTime, LocalTime time) {
      super(List.of(dateTime.toString(), time.toString()));
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"dateTime", "time"};
    }
  }

  static class MultiHeaderBean extends StringExcelTableBean {
    @ExcelColumn({"#", "#"}) @Nullable Integer rowNum;
    @ExcelColumn({"PersonalInfo", "Name"}) @Nullable String name;
    @ExcelColumn({"PersonalInfo", "Age"}) @Nullable Integer age;

    MultiHeaderBean(List<String> colList) {
      super(colList);
    }
  }

  // --- helpers ---

  private static void setCell(Sheet sheet, int poiRow, int poiCol, @Nullable String value) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      row = sheet.createRow(poiRow);
    }
    if (value == null) {
      row.createCell(poiCol);
    } else {
      row.createCell(poiCol).setCellValue(value);
    }
  }

  private Path buildTemplate(Workbook wb) throws Exception {
    Path file = tempDir.resolve("template.xlsx");
    try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
      wb.write(fos);
    }
    return file;
  }

  // --- tests ---

  @Nested
  @DisplayName("Field-order mapping (getFieldNameArray)")
  class FieldNameArrayMapping {

    @Test
    @DisplayName("writeFromBean → fields are written to cells in field order")
    void writesFieldsInOrder() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<FieldOrderBean>(
          "Sheet1", new String[] {"name", "age"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(), List.of(
          new FieldOrderBean(List.of("Alice", "25")),
          new FieldOrderBean(List.of("Bob", "30"))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Alice");
        assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("25");
        assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("Bob");
        assertThat(s.getRow(2).getCell(1).getStringCellValue()).isEqualTo("30");
      }
    }
  }

  @Nested
  @DisplayName("@ExcelColumn annotation")
  class ExcelColumnAnnotation {

    @Test
    @DisplayName("@ExcelColumn label match → maps to the correct field")
    void annotatedBeanMappedByLabel() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<AnnotatedBean>(
          "Sheet1", new String[] {"name", "age"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new AnnotatedBean(List.of("Alice", "25"))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Alice");
        assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("25");
      }
    }

    @Test
    @DisplayName("even when Excel column order is reversed from the header, @ExcelColumn maps by label")
    void columnOrderIndependent() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "age");   // reversed: age before name
        setCell(sheet, 0, 1, "name");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<AnnotatedBean>(
          "Sheet1", new String[] {"age", "name"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new AnnotatedBean(List.of("Alice", "25"))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("25");    // age col
        assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Alice"); // name col
      }
    }

    @Test
    @DisplayName("superclass @ExcelColumn fields are also inherited and mapped")
    void inheritedAnnotation() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "id");
        setCell(sheet, 0, 1, "name");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<AnnotatedSubBean>(
          "Sheet1", new String[] {"id", "name"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new AnnotatedSubBean(List.of("1", "Alice"))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("1");
        assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("Alice");
      }
    }
  }

  @Nested
  @DisplayName("Date/time formatting")
  class DateTimeFormatting {

    @Test
    @DisplayName("LocalDate → written in default ISO format (yyyy-MM-dd)")
    void localDateDefaultFormat() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "date");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<DateBean>(
          "Sheet1", new String[] {"date"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateBean(LocalDate.of(2026, 1, 15))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        assertThat(out.getSheet("Sheet1").getRow(1).getCell(0).getStringCellValue())
            .isEqualTo("2026-01-15");
      }
    }

    @Test
    @DisplayName("defaultDateTimeFormat specified → written in custom format")
    void localDateCustomFormat() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "date");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<DateBean>(
          "Sheet1", new String[] {"date"})
          .defaultDateTimeFormat(DateTimeFormatter.ofPattern("yyyy/MM/dd"));
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateBean(LocalDate.of(2026, 1, 15))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        assertThat(out.getSheet("Sheet1").getRow(1).getCell(0).getStringCellValue())
            .isEqualTo("2026/01/15");
      }
    }

    @Test
    @DisplayName("LocalDateTime is written via dateTimeFormatter, LocalTime via toString()")
    void localDateTimeAndLocalTime() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "dateTime");
        setCell(sheet, 0, 1, "time");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      LocalDateTime dt = LocalDateTime.of(2026, 1, 15, 9, 30, 0);
      LocalTime time = LocalTime.of(9, 30, 0);
      var writer = new StringOneLineHeaderExcelTableFromBeanWriter<DateTimeBean>(
          "Sheet1", new String[] {"dateTime", "time"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateTimeBean(dt, time)));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        // LocalDateTime uses dateTimeFormatter (ISO_LOCAL_DATE by default → date part only)
        assertThat(s.getRow(1).getCell(0).getStringCellValue()).isEqualTo("2026-01-15");
        assertThat(s.getRow(1).getCell(1).getStringCellValue()).isEqualTo("09:30");
      }
    }
  }

  @Nested
  @DisplayName("Multi-row header (StringHeaderExcelTableFromBeanWriter)")
  class MultiRowHeader {

    @Test
    @DisplayName("two-row header with @ExcelColumn({group, col}) → maps correctly and writes")
    void twoRowHeaderWithExcelColumn() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "PersonalInfo");
        setCell(sheet, 0, 2, "PersonalInfo");
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "Name");
        setCell(sheet, 1, 2, "Age");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new StringHeaderExcelTableFromBeanWriter<MultiHeaderBean>("Sheet1",
          new String[][] {{"#", "PersonalInfo", "PersonalInfo"}, {"#", "Name", "Age"}});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new MultiHeaderBean(List.of("1", "Alice", "25"))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        assertThat(s.getRow(2).getCell(0).getStringCellValue()).isEqualTo("1");
        assertThat(s.getRow(2).getCell(1).getStringCellValue()).isEqualTo("Alice");
        assertThat(s.getRow(2).getCell(2).getStringCellValue()).isEqualTo("25");
      }
    }
  }
}
