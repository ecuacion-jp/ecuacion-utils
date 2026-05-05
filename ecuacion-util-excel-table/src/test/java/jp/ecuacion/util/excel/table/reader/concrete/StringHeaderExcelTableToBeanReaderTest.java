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

@DisplayName("StringHeaderExcelTableToBeanReader")
public class StringHeaderExcelTableToBeanReaderTest {

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

  /** @ExcelColumn を使う Bean（getFieldNameArray オーバーライドなし）。 */
  static class AnnotatedBean extends StringExcelTableBean {
    @ExcelColumn("name") @NotBlank @Nullable String name;
    @ExcelColumn("age") @Min(1) @Nullable Integer age;

    public AnnotatedBean(List<String> colList) {
      super(colList);
    }
  }

  /** @ExcelColumn を持つスーパークラス。 */
  static class AnnotatedBaseBean extends StringExcelTableBean {
    @ExcelColumn("id") @Min(1) @Nullable Integer id;

    public AnnotatedBaseBean(List<String> colList) {
      super(colList);
    }
  }

  /** スーパークラスと自クラス両方に @ExcelColumn を持つ Bean。 */
  static class AnnotatedSubBean extends AnnotatedBaseBean {
    @ExcelColumn("name") @Nullable String name;

    public AnnotatedSubBean(List<String> colList) {
      super(colList);
    }
  }

  /** @ExcelColumn なし・getFieldNameArray オーバーライドなし。 */
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
  @DisplayName("正常系")
  class Normal {

    @Test
    @DisplayName("バリデーション違反なし → List を正常返却")
    void noViolations() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
        List<TestBean> result = reader.readToBean(file.toString());

        assertThat(result).hasSize(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }

    @Test
    @DisplayName("validates=false → バリデーション違反があっても例外なし")
    void validatesFalse() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
        assertThatNoException().isThrownBy(() -> reader.readToBean(file.toString(), false));
      }
    }
  }

  @Nested
  @DisplayName("行番号付きエラーメッセージ")
  class RowNumberInMessage {

    private @Nullable Arg getPostfix(Throwable ex) {
      Violations violations = ((ViolationException) ex).getViolations();
      return violations.messageParameters().getMessagePostfix();
    }

    @Test
    @DisplayName("データ1行目に違反 → messagePostfix の行番号引数が 2")
    void firstDataRowViolation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // Excel row 2, @Min(1) violation
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat(postfix.getMessageArgs()[1].getArgString()).isEqualTo("2");
            });
      }
    }

    @Test
    @DisplayName("データ2行目に違反 → messagePostfix の行番号引数が 3")
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

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat(postfix.getMessageArgs()[1].getArgString()).isEqualTo("3");
            });
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=3 → データ1行目の違反で行番号引数が 4")
    void tableStartRowNumber3() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Header at Excel row 3 (POI index 2)
        setCell(sheet, 2, 0, "name");
        setCell(sheet, 2, 1, "age");
        setCell(sheet, 3, 0, "Alice");
        setCell(sheet, 3, 1, "-1"); // Excel row 4
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 3, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat(postfix.getMessageArgs()[1].getArgString()).isEqualTo("4");
            });
      }
    }

    @Test
    @DisplayName("tableStartRowNumber=null（自動検索）→ 実際の行番号引数が含まれる")
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

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, null, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat(postfix.getMessageArgs()[1].getArgString()).isEqualTo("3");
            });
      }
    }

    @Test
    @DisplayName("messagePostfix にシート名引数が含まれる")
    void sheetNameInMessage() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("EmployeeSheet");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "EmployeeSheet", new String[] {"name", "age"}, 1, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString()))
            .isInstanceOf(ViolationException.class)
            .satisfies(ex -> {
              Arg postfix = Objects.requireNonNull(getPostfix(ex));
              assertThat(postfix.getMessageArgs()[0].getArgString()).isEqualTo("EmployeeSheet");
            });
      }
    }
  }

  @Nested
  @DisplayName("afterReading()")
  class AfterReading {

    @Test
    @DisplayName("afterReading() が RuntimeException をスローした場合伝播する")
    void afterReadingExceptionPropagates() {
      var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
          TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null) {
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
      };

      assertThatThrownBy(() -> reader.readToBean("dummy"))
          .isInstanceOf(RuntimeException.class)
          .hasMessage("afterReading error");
    }
  }

  @Nested
  @DisplayName("@ExcelColumn アノテーション")
  class ExcelColumnAnnotation {

    @Test
    @DisplayName("@ExcelColumn のラベルが headerLabels と同順 → 正しくマッピング")
    void sameOrderAsHeaderLabels() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result).hasSize(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("列順独立: headerLabels と異なる列順でも @ExcelColumn で正しくマッピング")
    void columnOrderIndependent() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // Excel の列順は age → name（Bean の宣言順 name → age と逆）
        setCell(sheet, 0, 0, "age");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "25");
        setCell(sheet, 1, 1, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"age", "name"}, 1, 1, null);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("headerLabels に対応する @ExcelColumn がない列はスキップ")
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

        // AnnotatedBean には @ExcelColumn("memo") がないので memo 列はスキップされる
        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "memo", "age"}, 1, 1, null);
        List<AnnotatedBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("スーパークラスの @ExcelColumn フィールドも有効")
    void inheritedAnnotations() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "id");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "1");
        setCell(sheet, 1, 1, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedSubBean>(
            AnnotatedSubBean.class, "Sheet1", new String[] {"id", "name"}, 1, 1, null);
        List<AnnotatedSubBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).id).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }

    @Test
    @DisplayName("@ExcelColumn のラベルが headerLabels に存在しない → RuntimeException")
    void unknownColumnLabel() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // headerLabels に "age" がない（"name" だけ）
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 1, 0, "Alice");
        Path file = writeTempExcel(wb);

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name"}, 1, 1, null);

        assertThatThrownBy(() -> reader.readToBean(file.toString(), false))
            .isInstanceOf(RuntimeException.class)
            .hasMessageContaining("age");
      }
    }

    @Test
    @DisplayName("getFieldNameArray() オーバーライドあり・@ExcelColumn なし → 従来通り動く")
    void backwardCompatibleOverride() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "25");
        Path file = writeTempExcel(wb);

        // TestBean は getFieldNameArray() をオーバーライドしており @ExcelColumn は使わない
        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
        List<TestBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("@ExcelColumn なし・getFieldNameArray() もなし → RuntimeException")
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
    @DisplayName("@ExcelColumn bean の1フィールド違反 → そのセルのみ赤くなる")
    void singleFieldViolation() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1"); // @Min(1) 違反
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
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
    @DisplayName("@ExcelColumn bean の複数フィールド違反 → 複数セルが赤くなる")
    void multipleFieldViolations() throws Exception {
      // @NotBlank + @Min(1) を持つ Bean を直接使う
      // name=null(@NotBlank 違反) + age=-1(@Min(1) 違反)
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, null);  // @NotBlank 違反
        setCell(sheet, 1, 1, "-1"); // @Min(1) 違反
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
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
    @DisplayName("tableStartColumnNumber=2 → 列オフセットが考慮されて正しい列がハイライト")
    void respectsTableStartColumnNumber() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        // テーブルが B 列(POI col=1)から始まる
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 0, 2, "age");
        setCell(sheet, 1, 1, "Alice");
        setCell(sheet, 1, 2, "-1"); // age 違反
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedBean>(
            AnnotatedBean.class, "Sheet1", new String[] {"name", "age"}, 1, 2, null);
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
    @DisplayName("スーパークラスの @ExcelColumn フィールド違反 → 継承チェーンを辿ってハイライト")
    void inheritedFieldHighlighted() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "id");
        setCell(sheet, 0, 1, "name");
        setCell(sheet, 1, 0, "-1"); // @Min(1) 違反（スーパークラスのフィールド）
        setCell(sheet, 1, 1, "Alice");
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        var reader = new StringHeaderExcelTableToBeanReader<AnnotatedSubBean>(
            AnnotatedSubBean.class, "Sheet1", new String[] {"id", "name"}, 1, 1, null);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          assertThat(isRedCell(s, 1, 0)).isTrue();  // id (スーパークラス) → red
          assertThat(isRedCell(s, 1, 1)).isFalse(); // name → not red
        }
      }
    }

    @Test
    @DisplayName("@ExcelColumn なし bean → 違反行の全データセルが赤くなる")
    void nonAnnotatedBeanHighlightsEntireRow() throws Exception {
      try (Workbook wb = new XSSFWorkbook()) {
        Sheet sheet = wb.createSheet("Sheet1");
        setCell(sheet, 0, 0, "name");
        setCell(sheet, 0, 1, "age");
        setCell(sheet, 1, 0, "Alice");
        setCell(sheet, 1, 1, "-1");
        Path input = writeTempExcel(wb);
        Path output = tempDir.resolve("output.xlsx");

        // TestBean は getFieldNameArray() override あり、@ExcelColumn なし
        var reader = new StringHeaderExcelTableToBeanReader<TestBean>(
            TestBean.class, "Sheet1", new String[] {"name", "age"}, 1, 1, null);
        try {
          reader.readToBean(input.toString());
        } catch (ViolationException ex) {
          reader.highlightErrors(input.toString(), ex.getViolations(), output.toString());
        }

        try (Workbook out = ExcelReadUtil.openForRead(output.toString())) {
          Sheet s = out.getSheet("Sheet1");
          // 全データセルが赤くなる
          assertThat(isRedCell(s, 1, 0)).isTrue(); // name → red
          assertThat(isRedCell(s, 1, 1)).isTrue(); // age  → red
        }
      }
    }
  }

  @Nested
  @DisplayName("複数行ヘッダー + @ExcelColumn")
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
    @DisplayName("2行ヘッダーで @ExcelColumn({\"group\",\"col\"}) が正しくマッピングされる")
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
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result).hasSize(1);
        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("列順独立: Excel の列順が変わっても @ExcelColumn で正しくマッピング")
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
            new String[][] {{"個人情報", "個人情報", "#"}, {"年齢", "名前", "#"}},
            1, 1, null);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
        assertThat(result.get(0).age).isEqualTo(25);
      }
    }

    @Test
    @DisplayName("@ExcelColumn(\"#\") 単一要素が縦結合列（全行一致）にマッチする")
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
            new String[][] {{"#", "個人情報", "個人情報"}, {"#", "名前", "年齢"}},
            1, 1, null);
        List<MultiHeaderBean> result = reader.readToBean(file.toString(), false);

        assertThat(result.get(0).rowNum).isEqualTo(1);
        assertThat(result.get(0).name).isEqualTo("Alice");
      }
    }
  }
}
