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
import static org.assertj.core.api.Assertions.assertThatNoException;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import java.io.FileOutputStream;
import java.util.Objects;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.lib.core.util.PropertiesFileUtil.Arg;
import jp.ecuacion.lib.core.violation.Violations;
import jp.ecuacion.util.excel.table.bean.ExcelColumn;
import jp.ecuacion.util.excel.table.bean.StringExcelTableBean;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("StringOneLineHeaderExcelTableToBeanReader / StringHeaderExcelTableToBeanReader")
public class StringHeaderExcelTableToBeanReaderTest {

  @SuppressWarnings("null")
  @TempDir
  Path tempDir;

  /** Test bean: name (@NotBlank), age (@Min(1)), with getFieldNameArray override. */
  static class TestBean extends StringExcelTableBean {
    @NotBlank @Nullable String name;
    @Min(1) @Nullable Integer age;

    public TestBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected String[] getFieldNameArray() {
      return new String[] {"name", "age"};
    }
  }

  /** Bean using {@code @ExcelColumn} (no getFieldNameArray override). */
  static class AnnotatedBean extends StringExcelTableBean {
    @ExcelColumn("name") @NotBlank @Nullable String name;
    @ExcelColumn("age") @Min(1) @Nullable Integer age;

    public AnnotatedBean(List<String> colList) {
      super(colList);
    }
  }

  /** Superclass with {@code @ExcelColumn}. */
  static class AnnotatedBaseBean extends StringExcelTableBean {
    @ExcelColumn("id") @Min(1) @Nullable Integer id;

    public AnnotatedBaseBean(List<String> colList) {
      super(colList);
    }
  }

  /** Bean with @ExcelColumn on both superclass and subclass. */
  static class AnnotatedSubBean extends AnnotatedBaseBean {
    @ExcelColumn("name") @Nullable String name;

    public AnnotatedSubBean(List<String> colList) {
      super(colList);
    }
  }

  /** Bean with no {@code @ExcelColumn} and no getFieldNameArray override. */
  static class NoAnnotationNoOverrideBean extends StringExcelTableBean {
    @Nullable String name;

    public NoAnnotationNoOverrideBean(List<String> colList) {
      super(colList);
    }
  }

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

  private Path writeTempExcel(Workbook wb) throws Exception {
    Path file = tempDir.resolve("test.xlsx");
    try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
      wb.write(fos);
    }
    return file;
  }

  @Nested
  @DisplayName("Normal")
  class Normal {

    @Test
    @DisplayName("no violations → returns list successfully")
    void noViolations() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        List<TestBean> result = reader.readToBean(file.toString());

        assertThat(result).hasSize(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }

    @Test
    @DisplayName("validates=false → no exception even when violations exist")
    void validatesFalse() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        assertThatNoException().isThrownBy(() -> reader.readToBean(file.toString(), false));
      }
    }
  }

  @Nested
  @DisplayName("row number in error message")
  class RowNumberInMessage {

    private @Nullable Arg getPostfix(Throwable ex) {
      Violations violations = ((ViolationException) ex).getViolations();
      return violations.messageParameters().getMessagePostfix();
    }

    @Test
    @DisplayName("violation in first data row → messagePostfix row number arg is 2")
    void firstDataRowViolation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // Excel row 2, @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat((String) postfix.getMessageArgs()[1]).isEqualTo("2");
            });
      }
    }

    @Test
    @DisplayName("violation in second data row → messagePostfix row number arg is 3")
    void secondDataRowViolation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25"); // valid
        setCell(sheet, 2, 0, "Bob");
        setCell(sheet, 2, 1, "-1"); // Excel row 3, @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat((String) postfix.getMessageArgs()[1]).isEqualTo("3");
            });
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=3 → violation in first data row gives row number arg 4")
    void tableStartRowNumber3() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Header at Excel row 3 (POI index 2)
        setCell(sheet, 2, 0, "name");
        setCell(sheet, 2, 1, "age");
        setCell(sheet, 3, 0, "Alice");
        setCell(sheet, 3, 1, "-1"); // Excel row 4
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(3);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat((String) postfix.getMessageArgs()[1]).isEqualTo("4");
            });
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=null (auto-detect) → actual row number is in the message")
    void autoDetectStartRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "unrelated");
        // Header at Excel row 2 (POI index 1)
        setCell(sheet, 1, 0, "name");
        setCell(sheet, 1, 1, "age");
        setCell(sheet, 2, 0, "Alice");
        setCell(sheet, 2, 1, "-1"); // Excel row 3
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"});

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat((String) postfix.getMessageArgs()[1]).isEqualTo("3");
            });
      }
    }

    @Test
    @DisplayName("sheet name is included in messagePostfix")
    void sheetNameInMessage() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("EmployeeSheet");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "EmployeeSheet", new String[] {"name", "age"})
            .tableStartRowNumber(1);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat((String) postfix.getMessageArgs()[0]).isEqualTo("EmployeeSheet");
            });
      }
    }
  }

  @Nested
  @DisplayName("afterReading()")
  class AfterReading {

    @Test
    @DisplayName("afterReading() RuntimeException propagates")
    void afterReadingExceptionPropagates() {
      var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
          TestBean.class, "Sheet1", new String[] {"name", "age"}) {
        @Override
        protected List<TestBean> excelTableToBeanList(String filePath) {
          TestBean bean = new TestBean(List.of("Alice", "25")) {
            @Override
            public void afterReading() {
              throw new RuntimeException("afterReading error");
            }
          };
          return List.of(bean);
        }
      }.tableStartRowNumber(1);

      assertThatThrownBy(() -> reader.readToBean("dummy"))
          .isInstanceOf(RuntimeException.class)
          .hasMessage("afterReading error");
    }
  }

  @Nested
  @DisplayName("@ExcelColumn annotation")
  class ExcelColumnAnnotation {

    @Test
    @DisplayName("@ExcelColumn labels in same order as headerLabels → mapped correctly")
    void sameOrderAsHeaderLabels() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result).hasSize(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("column-order independent: @ExcelColumn maps correctly regardless of column order")
    void columnOrderIndependent() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Excel column order is age → name (opposite of Bean field declaration order name → age)
        setCell(sheet, 0, 0, "age");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "25");
        setCell(sheet, 1, 1, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"age", "name"}).tableStartRowNumber(1);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("columns with no matching @ExcelColumn are skipped")
    void skipsUnmappedColumns() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "memo");
        setCell(sheet, 0, 2, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "some memo");
        setCell(sheet, 1, 2, "25");
        Path file = writeTempExcel(wb);

        // AnnotatedBean has no @ExcelColumn("memo"), so the "memo" column is skipped
        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "memo", "age"})
            .tableStartRowNumber(1);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("@ExcelColumn on superclass fields is also effective")
    void inheritedAnnotations() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "id");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "1");
        setCell(sheet, 1, 1, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedSubBean>(
            AnnotatedSubBean.class, "Sheet1", new String[] {"id", "name"}).tableStartRowNumber(1);
        List<AnnotatedSubBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).id).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }

    @Test
    @DisplayName("@ExcelColumn label not found in headerLabels → RuntimeException")
    void unknownColumnLabel() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // headerLabels has no "age" (only "name")
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 1, 0, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name"}).tableStartRowNumber(1);

        assertThatThrownBy(() -> reader.readToBean(file.toString(), false))
            .isInstanceOf(RuntimeException.class)
            .hasMessageContaining("age");
      }
    }

    @Test
    @DisplayName("getFieldNameArray() override without @ExcelColumn → works as before")
    void backwardCompatibleOverride() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        // TestBean overrides getFieldNameArray() and does not use @ExcelColumn
        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        List<TestBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("no @ExcelColumn and no getFieldNameArray() → RuntimeException")
    void noAnnotationNoOverride() {
      assertThatThrownBy(() -> new NoAnnotationNoOverrideBean(List.of("Alice")))
          .isInstanceOf(RuntimeException.class)
          .hasMessageContaining("@ExcelColumn");
    }
  }

  @Nested
  @DisplayName("highlightErrors")
  class HighlightErrors {

    private boolean isRedCell(Sheet sheet, int poiRow, int poiCol) {
      Row row = sheet.getRow(poiRow);
      if (row == null) {
        return false;
      }
      var cell = row.getCell(poiCol);
      if (cell == null) {
        return false;
      }
      var style = cell.getCellStyle();
      return style.getFillPattern() == FillPatternType.SOLID_FOREGROUND
          && style.getFillForegroundColor() == IndexedColors.RED1.getIndex();
    }

    @Test
    @DisplayName("single field violation in @ExcelColumn bean → only that cell turns red")
    void singleFieldViolation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // @Min(1) violation
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          assertThat(isRedCell(s, 1, 1)).isTrue();   // age cell → red
          assertThat(isRedCell(s, 1, 0)).isFalse();  // name cell → not red
        }
      }
    }

    @Test
    @DisplayName("multiple field violations in @ExcelColumn bean → multiple cells turn red")
    void multipleFieldViolations() throws Exception {
      // Use a bean with both @NotBlank and @Min(1) directly
      // name=null (@NotBlank violation) + age=-1 (@Min(1) violation)
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, null);  // @NotBlank violation
        setCell(sheet, 1, 1, "-1"); // @Min(1) violation
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          assertThat(isRedCell(s, 1, 0)).isTrue(); // name → red
          assertThat(isRedCell(s, 1, 1)).isTrue(); // age  → red
        }
      }
    }

    @Test
    @DisplayName("tableStartColumnNumber=2 → column offset respected, correct column highlighted")
    void respectsTableStartColumnNumber() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Table starts at column B (POI col=1)
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 0, 2, "age");
        setCell(sheet, 1, 1, "Alice");
        setCell(sheet, 1, 2, "-1"); // age violation
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"})
            .tableStartRowNumber(1).tableStartColumnNumber(2);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          assertThat(isRedCell(s, 1, 2)).isTrue();  // age (POI col=2) → red
          assertThat(isRedCell(s, 1, 1)).isFalse(); // name (POI col=1) → not red
        }
      }
    }

    @Test
    @DisplayName("superclass @ExcelColumn field violation → highlighted by traversing inheritance chain")
    void inheritedFieldHighlighted() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "id");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "-1"); // @Min(1) violation (superclass field)
        setCell(sheet, 1, 1, "Alice");
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringOneLineHeaderExcelTableToBeanReader<AnnotatedSubBean>(
            AnnotatedSubBean.class, "Sheet1", new String[] {"id", "name"}).tableStartRowNumber(1);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          assertThat(isRedCell(s, 1, 0)).isTrue();  // id (superclass field) → red
          assertThat(isRedCell(s, 1, 1)).isFalse(); // name → not red
        }
      }
    }

    @Test
    @DisplayName("bean without @ExcelColumn → all data cells in the violation row turn red")
    void nonAnnotatedBeanHighlightsEntireRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1");
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        // TestBean has getFieldNameArray() override, no @ExcelColumn
        var reader = new StringOneLineHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}).tableStartRowNumber(1);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          // All data cells turn red
          assertThat(isRedCell(s, 1, 0)).isTrue(); // name → red
          assertThat(isRedCell(s, 1, 1)).isTrue(); // age  → red
        }
      }
    }
  }

  @Nested
  @DisplayName("multi-row header + @ExcelColumn")
  class MultiLineHeaderWithExcelColumn {

    /** Bean for 2-row header: group + column. */
    static class MultiHeaderBean extends StringExcelTableBean {
      @ExcelColumn({"#", "#"}) @Nullable Integer rowNum;
      @ExcelColumn({"個人情報", "名前"}) @Nullable String name;
      @ExcelColumn({"個人情報", "年齢"}) @Nullable Integer age;

      public MultiHeaderBean(List<String> colList) {
        super(colList);
      }
    }

    @Test
    @DisplayName("2-row header with @ExcelColumn({\"group\",\"col\"}) is mapped correctly")
    void multiRowHeaderMapping() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "個人情報");
        setCell(sheet, 1, 0, "#");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<MultiHeaderBean>(
            MultiHeaderBean.class, "Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}})
            .tableStartRowNumber(1);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result).hasSize(1);
        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("column-order independent: @ExcelColumn maps correctly even when Excel column order changes")
    void multiRowHeaderColumnOrderIndependent() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Excel order: age, name, # (reversed from bean declaration)
        setCell(sheet, 0, 0, "個人情報");
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "#");
        setCell(sheet, 1, 0, "年齢");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "#");
        setCell(sheet, 2, 0, "25");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "1");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<MultiHeaderBean>(
            MultiHeaderBean.class, "Sheet1",
            new String[][] {{"個人情報", "個人情報", "#"}, {"年齢", "名前", "#"}})
            .tableStartRowNumber(1);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("single-element @ExcelColumn(\"#\") matches a vertically merged column")
    void singleElementAnnotationMatchesVerticalMerge() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "#");
        setCell(sheet, 0, 1, "個人情報");
        setCell(sheet, 0, 2, "個人情報");
        setCell(sheet, 1, 1, "名前");
        setCell(sheet, 1, 2, "年齢");
        setCell(sheet, 2, 0, "1");
        setCell(sheet, 2, 1, "Alice");
        setCell(sheet, 2, 2, "25");
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<MultiHeaderBean>(
            MultiHeaderBean.class, "Sheet1",
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}})
            .tableStartRowNumber(1);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }
  }
}
