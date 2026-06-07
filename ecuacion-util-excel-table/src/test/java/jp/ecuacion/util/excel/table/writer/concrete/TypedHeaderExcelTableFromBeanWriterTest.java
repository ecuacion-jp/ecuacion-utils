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
import java.util.List;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.TypedExcelTableBean;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("TypedOneLineHeaderExcelTableFromBeanWriter / TypedHeaderExcelTableFromBeanWriter")
public class TypedHeaderExcelTableFromBeanWriterTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  // --- test beans ---

  static class PersonBean extends TypedExcelTableBean {
    @Nullable String name;
    @Nullable Integer age;
    @Nullable Double score;
    @Nullable Boolean active;

    public PersonBean(List<Object> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"name", "age", "score", "active"};
    }
  }

  static class AnnotatedBean extends TypedExcelTableBean {
    @ExcelColumn("name") @Nullable String name;
    @ExcelColumn("age") @Nullable Integer age;

    public AnnotatedBean(List<Object> colList) {
      super(colList);
    }
  }

  static class DateBean extends TypedExcelTableBean {
    @Nullable LocalDate date;

    public DateBean(LocalDate date) {
      super(List.of(date));
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"date"};
    }
  }

  static class DateTimeBean extends TypedExcelTableBean {
    @Nullable LocalDateTime startsAt;

    public DateTimeBean(LocalDateTime startsAt) {
      super(List.of(startsAt));
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"startsAt"};
    }
  }

  // --- helpers ---

  private static void setStringCell(Sheet sheet, int poiRow, int poiCol, String value) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      row = sheet.createRow(poiRow);
    }
    row.createCell(poiCol).setCellValue(value);
  }

  /** Pre-formats the cell at {@code (poiRow, poiCol)} with a custom date format pattern. */
  private static void preFormatAsDate(Workbook wb, Sheet sheet, int poiRow, int poiCol,
      String formatPattern) {
    Row row = sheet.getRow(poiRow);
    if (row == null) {
      row = sheet.createRow(poiRow);
    }
    Cell cell = row.createCell(poiCol);
    CellStyle style = wb.createCellStyle();
    style.setDataFormat(wb.createDataFormat().getFormat(formatPattern));
    cell.setCellStyle(style);
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
  @DisplayName("フィールド順マッピング (getFieldNameArray)")
  class FieldNameArrayMapping {

    @Test
    @DisplayName("writeFromBean → 各フィールドがネイティブ型のままセルへ書き込まれる")
    void writesNativeTypedValues() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "name");
        setStringCell(sheet, 0, 1, "age");
        setStringCell(sheet, 0, 2, "score");
        setStringCell(sheet, 0, 3, "active");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<PersonBean>("Sheet1",
          new String[] {"name", "age", "score", "active"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new PersonBean(List.of("Alice", 25.0, 92.5, true))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        Row row = s.getRow(1);
        assertThat(row.getCell(0).getStringCellValue()).isEqualTo("Alice");
        assertThat(row.getCell(1).getNumericCellValue()).isEqualTo(25.0);
        assertThat(row.getCell(2).getNumericCellValue()).isEqualTo(92.5);
        assertThat(row.getCell(3).getBooleanCellValue()).isTrue();
      }
    }
  }

  @Nested
  @DisplayName("@ExcelColumn アノテーション")
  class ExcelColumnAnnotation {

    @Test
    @DisplayName("Excel の列順がヘッダーと逆でも @ExcelColumn でラベル一致マッピングして書き込まれる")
    void columnOrderIndependent() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "age");
        setStringCell(sheet, 0, 1, "name");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<AnnotatedBean>("Sheet1",
          new String[] {"age", "name"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new AnnotatedBean(List.of("Alice", 25.0))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        Row row = s.getRow(1);
        assertThat(row.getCell(0).getNumericCellValue()).isEqualTo(25.0); // age col
        assertThat(row.getCell(1).getStringCellValue()).isEqualTo("Alice"); // name col
      }
    }
  }

  @Nested
  @DisplayName("日付セルの書式保証")
  class DateCellFormatting {

    @Test
    @DisplayName("テンプレートのセルに日付書式が無い → デフォルト書式 (yyyy-mm-dd) が適用され日付として書き込まれる")
    void appliesDefaultDateFormatWhenAbsent() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "date");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<DateBean>("Sheet1",
          new String[] {"date"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateBean(LocalDate.of(2026, 1, 15))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Cell cell = out.getSheet("Sheet1").getRow(1).getCell(0);
        assertThat(DateUtil.isCellDateFormatted(cell)).isTrue();
        assertThat(cell.getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd");
        assertThat(cell.getLocalDateTimeCellValue().toLocalDate()).isEqualTo(LocalDate.of(2026, 1, 15));
      }
    }

    @Test
    @DisplayName("テンプレートのセルに既に日付書式がある → その書式を変更せず値のみ書き込まれる")
    void keepsExistingDateFormat() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "date");
        // Pre-format the data cell with a custom (non-default) date format.
        preFormatAsDate(wb, sheet, 1, 0, "yyyy/mm/dd");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<DateBean>("Sheet1",
          new String[] {"date"});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateBean(LocalDate.of(2026, 1, 15))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Cell cell = out.getSheet("Sheet1").getRow(1).getCell(0);
        assertThat(DateUtil.isCellDateFormatted(cell)).isTrue();
        // The template's custom format is preserved, not overwritten by the default.
        assertThat(cell.getCellStyle().getDataFormatString()).isEqualTo("yyyy/mm/dd");
        assertThat(cell.getLocalDateTimeCellValue().toLocalDate()).isEqualTo(LocalDate.of(2026, 1, 15));
      }
    }

    @Test
    @DisplayName("LocalDateTime かつ書式なし → デフォルト書式 (yyyy-mm-dd hh:mm:ss) が適用される")
    void appliesDefaultDateTimeFormatWhenAbsent() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "startsAt");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<DateTimeBean>("Sheet1",
          new String[] {"startsAt"});
      writer.tableStartRowNumber(1);
      LocalDateTime dt = LocalDateTime.of(2026, 1, 15, 9, 30);
      writer.writeFromBean(template.toString(), output.toString(), List.of(new DateTimeBean(dt)));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Cell cell = out.getSheet("Sheet1").getRow(1).getCell(0);
        assertThat(DateUtil.isCellDateFormatted(cell)).isTrue();
        assertThat(cell.getCellStyle().getDataFormatString()).isEqualTo("yyyy-mm-dd hh:mm:ss");
        assertThat(cell.getLocalDateTimeCellValue()).isEqualTo(dt);
      }
    }

    @Test
    @DisplayName("defaultDateFormat 指定 → 書式の無いセルにはカスタム書式が適用される")
    void customDefaultDateFormat() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "date");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedOneLineHeaderExcelTableFromBeanWriter<DateBean>("Sheet1",
          new String[] {"date"}).defaultDateFormat("yyyy/mm/dd");
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new DateBean(LocalDate.of(2026, 1, 15))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Cell cell = out.getSheet("Sheet1").getRow(1).getCell(0);
        assertThat(cell.getCellStyle().getDataFormatString()).isEqualTo("yyyy/mm/dd");
        assertThat(cell.getLocalDateTimeCellValue().toLocalDate()).isEqualTo(LocalDate.of(2026, 1, 15));
      }
    }
  }

  @Nested
  @DisplayName("複数行ヘッダー (TypedHeaderExcelTableFromBeanWriter)")
  class MultiRowHeader {

    static class MultiHeaderBean extends TypedExcelTableBean {
      @ExcelColumn({"#", "#"}) @Nullable Integer rowNum;
      @ExcelColumn({"個人情報", "名前"}) @Nullable String name;
      @ExcelColumn({"個人情報", "年齢"}) @Nullable Integer age;

      public MultiHeaderBean(List<Object> colList) {
        super(colList);
      }
    }

    @Test
    @DisplayName("2行ヘッダーと @ExcelColumn({group, col}) → 正しくマッピングして書き込まれる")
    void twoRowHeaderWithExcelColumn() throws Exception {
      Path template;
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setStringCell(sheet, 0, 0, "#");
        setStringCell(sheet, 0, 1, "個人情報");
        setStringCell(sheet, 0, 2, "個人情報");
        setStringCell(sheet, 1, 0, "#");
        setStringCell(sheet, 1, 1, "名前");
        setStringCell(sheet, 1, 2, "年齢");
        template = buildTemplate(wb);
      }
      Path output = tempDir.resolve("output.xlsx");

      var writer = new TypedHeaderExcelTableFromBeanWriter<MultiHeaderBean>("Sheet1",
          new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}});
      writer.tableStartRowNumber(1);
      writer.writeFromBean(template.toString(), output.toString(),
          List.of(new MultiHeaderBean(List.of(1.0, "Alice", 25.0))));

      try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
        Sheet s = out.getSheet("Sheet1");
        Row row = s.getRow(2);
        assertThat(row.getCell(0).getNumericCellValue()).isEqualTo(1.0);
        assertThat(row.getCell(1).getStringCellValue()).isEqualTo("Alice");
        assertThat(row.getCell(2).getNumericCellValue()).isEqualTo(25.0);
      }
    }
  }
}
