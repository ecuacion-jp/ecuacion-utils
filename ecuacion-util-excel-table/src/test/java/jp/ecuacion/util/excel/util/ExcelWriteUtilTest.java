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
package jp.ecuacion.util.excel.util;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatCode;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.io.File;
import java.time.LocalDate;
import java.util.stream.Stream;
import jp.ecuacion.lib.core.exception.ViolationException;
import jp.ecuacion.util.excel.exception.ExcelAppException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

@DisplayName("ExcelWriteUtil")
public class ExcelWriteUtilTest {

  private static final String EXCEL_PATH =
      new File("src/test/resources").getAbsolutePath() + "/ExcelWriteUtilTest.xlsx";

  @Nested
  @DisplayName("createWorkbookWithSheet()")
  class CreateWorkbookWithSheet {

    @Test
    @DisplayName("指定した名前のシートを持つ Workbook を返す")
    void createsWorkbookWithNamedSheet() throws Exception {
      try (Workbook wb = ExcelWriteUtil.createWorkbookWithSheet("MySheet")) {
        assertThat(wb.getSheet("MySheet")).isNotNull();
      }
    }
  }

  @Nested
  @DisplayName("getReadyToEvaluateFormula()")
  class GetReadyToEvaluateFormula {

    @Nested
    @DisplayName("STRING でないセルのとき")
    class WhenCellTypeIsNotString {

      @Test
      @DisplayName("変更なし")
      void unchanged() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(1.0);
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, true, false, false, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue()).isEqualTo(1.0);
        }
      }
    }

    @Nested
    @DisplayName("STRING セル × changesNumberString")
    class WhenChangesNumberString {

      @Test
      @DisplayName("changesNumberString=false → STRING のまま")
      void staysStringWhenFlagIsFalse() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("1");
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, false, false, false, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
          assertThat(cell.getStringCellValue()).isEqualTo("1");
        }
      }

      @ParameterizedTest(name = "[{index}] value={0} → NUMERIC {1}")
      @MethodSource
      @DisplayName("changesNumberString=true → NUMERIC に変換される")
      void convertsToNumeric(String value, double expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, true, false, false, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue()).isEqualTo(expected);
        }
      }

      static Stream<Arguments> convertsToNumeric() {
        return Stream.of(
            Arguments.of("1", 1.0),
            Arguments.of("1,234", 1234.0),
            Arguments.of("1.5", 1.5),
            Arguments.of("-1", -1.0));
      }

      @Test
      @DisplayName("changesNumberString=true、数値でない文字列 → STRING のまま")
      void staysStringWhenNotParseable() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("abc");
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, true, false, false, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
          assertThat(cell.getStringCellValue()).isEqualTo("abc");
        }
      }
    }

    @Nested
    @DisplayName("STRING セル × changesDateString")
    class WhenChangesDateString {

      @Test
      @DisplayName("changesDateString=false → STRING のまま")
      void staysStringWhenFlagIsFalse() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("2025/01/01");
          ExcelWriteUtil.getReadyToEvaluateFormula(
              cell, false, false, false, new String[]{"yyyy/MM/dd"});
          assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
          assertThat(cell.getStringCellValue()).isEqualTo("2025/01/01");
        }
      }

      @ParameterizedTest(name = "[{index}] formats={1}")
      @MethodSource
      @DisplayName("changesDateString=true → NUMERIC（日付シリアル値）に変換される")
      void convertsToDateSerial(String value, String[] formats) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, false, true, false, formats);
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue())
              .isEqualTo(DateUtil.getExcelDate(LocalDate.of(2025, 1, 1)));
        }
      }

      static Stream<Arguments> convertsToDateSerial() {
        return Stream.of(
            Arguments.of("2025/01/01", new String[]{"yyyy/MM/dd"}),
            Arguments.of("2025/01/01", new String[]{"yyyy-MM-dd", "yyyy/MM/dd"}));
      }

      @Test
      @DisplayName("changesDateString=true、フォーマット不一致 → STRING のまま")
      void staysStringWhenNoMatch() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("abc");
          ExcelWriteUtil.getReadyToEvaluateFormula(
              cell, false, true, false, new String[]{"yyyy/MM/dd"});
          assertThat(cell.getCellType()).isEqualTo(CellType.STRING);
          assertThat(cell.getStringCellValue()).isEqualTo("abc");
        }
      }
    }

    @Nested
    @DisplayName("テキストフォーマット（format==49）のとき")
    class WhenTextDataFormat {

      @ParameterizedTest(name = "[{index}] changesCellsWithTextDataFormat={0} → {1}")
      @MethodSource
      @DisplayName("changesCellsWithTextDataFormat に応じた動作")
      void behavior(boolean changesCells, CellType expectedType) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("1");
          CellStyle style = wb.createCellStyle();
          style.setDataFormat((short) 49);
          cell.setCellStyle(style);
          ExcelWriteUtil.getReadyToEvaluateFormula(
              cell, true, false, changesCells, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(expectedType);
        }
      }

      static Stream<Arguments> behavior() {
        return Stream.of(
            Arguments.of(false, CellType.STRING),
            Arguments.of(true, CellType.NUMERIC));
      }
    }

    @Nested
    @DisplayName("changesNumberString と changesDateString の両方 true のとき")
    class WhenBothFlagsTrue {

      @Test
      @DisplayName("value が数値文字列 → 数値変換が先に成功、日付変換はスキップされる")
      void numberStringConverts() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("1");
          ExcelWriteUtil.getReadyToEvaluateFormula(
              cell, true, true, false, new String[]{"yyyy/MM/dd"});
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue()).isEqualTo(1.0);
        }
      }

      @Test
      @DisplayName("value が日付文字列 → 数値変換が失敗、日付変換が成功する")
      void dateStringConverts() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue("2025/01/01");
          ExcelWriteUtil.getReadyToEvaluateFormula(
              cell, true, true, false, new String[]{"yyyy/MM/dd"});
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue())
              .isEqualTo(DateUtil.getExcelDate(LocalDate.of(2025, 1, 1)));
        }
      }
    }
  }

  @Nested
  @DisplayName("evaluateFormula()")
  class EvaluateFormula {

    @Nested
    @DisplayName("evaluateFormula(Cell, String)")
    class CellLevel {

      @Test
      @DisplayName("非数式セル（NUMERIC）→ 例外なし")
      void nonFormulaCell() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(123.0);
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      @Test
      @DisplayName("通常の数式 → 例外なし")
      void normalFormula() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellFormula("1+1");
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      @Test
      @DisplayName("未実装関数 → ExcelAppException（NotImplementedException が原因）")
      void unimplementedFunction() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(3).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExcelAppException.class)
              .extracting(e -> ((ExcelAppException) e).getMessageId())
              .isEqualTo(
                  "jp.ecuacion.util.excel.ExcelWriteUtil.NotImplementedException.message");
        }
      }

      @Test
      @DisplayName("外部ブック参照 → ExcelAppException（WorkbookNotFoundException が原因）")
      void externalWorkbookRef() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(5).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExcelAppException.class)
              .extracting(e -> ((ExcelAppException) e).getMessageId())
              .isEqualTo(
                  "jp.ecuacion.util.excel.ExcelWriteUtil.WorkbookNotFoundException.message");
        }
      }

      @Test
      @DisplayName("#NAME? → ExcelAppException（DetailUnknown、原因は FormulaParseException）")
      void namePound() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(4).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExcelAppException.class)
              .satisfies(e -> {
                assertThat(((ExcelAppException) e).getMessageId())
                    .isEqualTo(
                        "jp.ecuacion.util.excel.ExcelWriteUtil.DetailUnknown.message");
                assertThat(((ExcelAppException) e).getCause())
                    .isInstanceOf(FormulaParseException.class);
              });
        }
      }

      @ParameterizedTest(name = "[{index}] 行 {0} のエラー値セル → 例外なし")
      @MethodSource
      @DisplayName("エラー値（#VALUE! / #DIV/0! / #N/A）→ 例外なし")
      void errorValues(int rowIndex) throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(rowIndex).getCell(1);
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      static Stream<Arguments> errorValues() {
        return Stream.of(
            Arguments.of(6),
            Arguments.of(7),
            Arguments.of(8));
      }

      @Test
      @DisplayName("その他の例外 → ExcelAppException（DetailUnknown、原因は ClassCastException）")
      void otherException() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(9).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExcelAppException.class)
              .satisfies(e -> {
                assertThat(((ExcelAppException) e).getMessageId())
                    .isEqualTo(
                        "jp.ecuacion.util.excel.ExcelWriteUtil.DetailUnknown.message");
                assertThat(((ExcelAppException) e).getCause())
                    .isInstanceOf(ClassCastException.class);
              });
        }
      }
    }

    @Nested
    @DisplayName("evaluateFormula(Workbook, String, boolean)")
    class WorkbookLevel {

      @Test
      @DisplayName("breaksOnError=true → 最初のエラーで即 ExcelAppException")
      void breaksOnErrorTrue() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(wb, "file", true))
              .isInstanceOf(ExcelAppException.class);
        }
      }

      @Test
      @DisplayName("breaksOnError=false → 全エラーを収集して ViolationException")
      void breaksOnErrorFalse() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(wb, "file", false))
              .isInstanceOf(ViolationException.class)
              .satisfies(e -> assertThat(
                  ((ViolationException) e).getViolations().getBusinessViolations())
                  .hasSizeGreaterThan(1));
        }
      }
    }

    @Nested
    @DisplayName("evaluateFormula(Workbook, String, boolean, String...)")
    class WorkbookWithSheetsOverload {

      @Test
      @DisplayName("対象外シートのエラー数式は評価されない")
      void ignoresErrorsInNonTargetSheets() throws Exception {
        try (Workbook wb = ExcelWriteUtil.openForWrite(EXCEL_PATH)) {
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(
              wb, "file", false, "getReadyToEvaluateFormulaTest"))
              .doesNotThrowAnyException();
        }
      }
    }
  }
}
