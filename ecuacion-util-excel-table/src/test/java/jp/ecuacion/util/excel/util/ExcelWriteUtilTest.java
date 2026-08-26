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
import jp.ecuacion.util.excel.exception.ExcelFeatureNotImplementedException;
import jp.ecuacion.util.excel.exception.ExcelTableException;
import jp.ecuacion.util.excel.exception.ExternalWorkbookNotFoundException;
import jp.ecuacion.util.excel.exception.FormulaEvaluationUnknownErrorException;
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
import org.jspecify.annotations.Nullable;
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
    @DisplayName("returns a Workbook with a sheet of the given name")
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
    @DisplayName("when cell type is not STRING")
    class WhenCellTypeIsNotString {

      @Test
      @DisplayName("no change")
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
    @DisplayName("STRING cell × changesNumberString")
    class WhenChangesNumberString {

      @Test
      @DisplayName("changesNumberString=false → stays STRING")
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
      @DisplayName("changesNumberString=true → converted to NUMERIC")
      void convertsToNumeric(String value, double expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          ExcelWriteUtil.getReadyToEvaluateFormula(cell, true, false, false, new String[0]);
          assertThat(cell.getCellType()).isEqualTo(CellType.NUMERIC);
          assertThat(cell.getNumericCellValue()).isEqualTo(expected);
        }
      }

      static @Nullable Stream<@Nullable Arguments> convertsToNumeric() {
        return Stream.of(
            Arguments.of("1", 1.0),
            Arguments.of("1,234", 1234.0),
            Arguments.of("1.5", 1.5),
            Arguments.of("-1", -1.0));
      }

      @Test
      @DisplayName("changesNumberString=true, non-numeric string → stays STRING")
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
    @DisplayName("STRING cell × changesDateString")
    class WhenChangesDateString {

      @Test
      @DisplayName("changesDateString=false → stays STRING")
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
      @DisplayName("changesDateString=true → converted to NUMERIC (date serial value)")
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

      static @Nullable Stream<@Nullable Arguments> convertsToDateSerial() {
        return Stream.of(
            Arguments.of("2025/01/01", new String[]{"yyyy/MM/dd"}),
            Arguments.of("2025/01/01", new String[]{"yyyy-MM-dd", "yyyy/MM/dd"}));
      }

      @Test
      @DisplayName("changesDateString=true, no format match → stays STRING")
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
    @DisplayName("when using the text format (format==49)")
    class WhenTextDataFormat {

      @ParameterizedTest(name = "[{index}] changesCellsWithTextDataFormat={0} → {1}")
      @MethodSource
      @DisplayName("behavior depends on changesCellsWithTextDataFormat")
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

      static @Nullable Stream<@Nullable Arguments> behavior() {
        return Stream.of(
            Arguments.of(false, CellType.STRING),
            Arguments.of(true, CellType.NUMERIC));
      }
    }

    @Nested
    @DisplayName("when both changesNumberString and changesDateString are true")
    class WhenBothFlagsTrue {

      @Test
      @DisplayName("value is a numeric string → number conversion succeeds first, date conversion is skipped")
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
      @DisplayName("value is a date string → number conversion fails, date conversion succeeds")
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
      @DisplayName("non-formula cell (NUMERIC) → no exception")
      void nonFormulaCell() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(123.0);
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      @Test
      @DisplayName("normal formula → no exception")
      void normalFormula() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellFormula("1+1");
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      @Test
      @DisplayName("unimplemented function → ExcelFeatureNotImplementedException"
          + " (caused by NotImplementedException)")
      void unimplementedFunction() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(3).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExcelFeatureNotImplementedException.class);
        }
      }

      @Test
      @DisplayName("external workbook reference → ExternalWorkbookNotFoundException"
          + " (caused by WorkbookNotFoundException)")
      void externalWorkbookRef() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(5).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(ExternalWorkbookNotFoundException.class);
        }
      }

      @Test
      @DisplayName("#NAME? → FormulaEvaluationUnknownErrorException"
          + " (caused by FormulaParseException)")
      void namePound() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(4).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(FormulaEvaluationUnknownErrorException.class)
              .hasCauseInstanceOf(FormulaParseException.class);
        }
      }

      @ParameterizedTest(name = "[{index}] error value cell at row {0} → no exception")
      @MethodSource
      @DisplayName("error value (#VALUE! / #DIV/0! / #N/A) → no exception")
      void errorValues(int rowIndex) throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(rowIndex).getCell(1);
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .doesNotThrowAnyException();
        }
      }

      static @Nullable Stream<@Nullable Arguments> errorValues() {
        return Stream.of(
            Arguments.of(6),
            Arguments.of(7),
            Arguments.of(8));
      }

      @Test
      @DisplayName("other exceptions → FormulaEvaluationUnknownErrorException"
          + " (caused by ClassCastException)")
      void otherException() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          Cell cell = wb.getSheet("evaluateFormulaTest").getRow(9).getCell(1);
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(cell, "file"))
              .isInstanceOf(FormulaEvaluationUnknownErrorException.class)
              .hasCauseInstanceOf(ClassCastException.class);
        }
      }
    }

    @Nested
    @DisplayName("evaluateFormula(Workbook, String, boolean)")
    class WorkbookLevel {

      @Test
      @DisplayName("breaksOnError=true → immediately throws ExcelTableException on the first error")
      void breaksOnErrorTrue() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          assertThatThrownBy(() -> ExcelWriteUtil.evaluateFormula(wb, "file", true))
              .isInstanceOf(ExcelTableException.class);
        }
      }

      @Test
      @DisplayName("breaksOnError=false → collects all errors into a ViolationException")
      void breaksOnErrorFalse() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
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
      @DisplayName("error formulas in non-target sheets are not evaluated")
      void ignoresErrorsInNonTargetSheets() throws Exception {
        try (Workbook wb = ExcelReadUtil.openForRead(EXCEL_PATH)) {
          assertThatCode(() -> ExcelWriteUtil.evaluateFormula(
              wb, "file", false, "getReadyToEvaluateFormulaTest"))
              .doesNotThrowAnyException();
        }
      }
    }
  }
}
