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
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.stream.Stream;
import jp.ecuacion.util.excel.exception.ExcelAppException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

@DisplayName("ExcelReadUtil")
public class ExcelReadUtilTest {

  @Nested
  @DisplayName("getNoDataStringIfNoData()")
  class GetNoDataStringIfNoData {

    @Nested
    @DisplayName("value が null または空文字のとき")
    class WhenValueIsNullOrEmpty {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("noDataString をそのまま返す")
      void returnsNoDataString(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) {
        assertThat(ExcelReadUtil.getNoDataStringIfNoData(value, noDataString))
            .isEqualTo(expected);
      }

      static Stream<Arguments> returnsNoDataString() {
        return Stream.of(
            Arguments.of(null, null, null),
            Arguments.of(null, "", ""),
            Arguments.of(null, "N/A", "N/A"),
            Arguments.of("", null, null),
            Arguments.of("", "", ""),
            Arguments.of("", "N/A", "N/A"));
      }
    }

    @Nested
    @DisplayName("value が通常文字列のとき")
    class WhenValueIsNotEmpty {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("value をそのまま返す（noDataString は無視）")
      void returnsValue(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) {
        assertThat(ExcelReadUtil.getNoDataStringIfNoData(value, noDataString))
            .isEqualTo(expected);
      }

      static Stream<Arguments> returnsValue() {
        return Stream.of(
            Arguments.of("abc", null, "abc"),
            Arguments.of("abc", "N/A", "abc"),
            Arguments.of(" ", null, " "),
            Arguments.of("0", null, "0"));
      }
    }
  }

  @Nested
  @DisplayName("getStringFromCell()")
  class GetStringFromCell {

    @Nested
    @DisplayName("cell が null のとき")
    class WhenCellIsNull {

      @ParameterizedTest(name = "[{index}] noDataString={0} → {1}")
      @MethodSource
      @DisplayName("noDataString を返す")
      void returnsNoDataString(@Nullable String noDataString, @Nullable String expected)
          throws ExcelAppException {
        assertThat(ExcelReadUtil.getStringFromCell(null, null, null, noDataString))
            .isEqualTo(expected);
      }

      static Stream<Arguments> returnsNoDataString() {
        return Stream.of(
            Arguments.of(null, null),
            Arguments.of("N/A", "N/A"));
      }
    }

    @Nested
    @DisplayName("BLANK セルのとき")
    class WhenCellTypeIsBlank {

      @ParameterizedTest(name = "[{index}] noDataString={0} → {1}")
      @MethodSource
      @DisplayName("noDataString を返す")
      void returnsNoDataString(@Nullable String noDataString, @Nullable String expected)
          throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, null, noDataString))
              .isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsNoDataString() {
        return Stream.of(
            Arguments.of(null, null),
            Arguments.of("N/A", "N/A"));
      }
    }

    @Nested
    @DisplayName("STRING セルのとき")
    class WhenCellTypeIsString {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("文字列または noDataString を返す")
      void returnsExpected(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, null, noDataString))
              .isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of("hello", null, "hello"),
            Arguments.of("", null, null),
            Arguments.of("", "N/A", "N/A"),
            Arguments.of(" ", null, " "));
      }
    }

    @Nested
    @DisplayName("NUMERIC セル（表示形式：標準）のとき")
    class WhenCellTypeIsNumericWithGeneralFormat {

      @ParameterizedTest(name = "[{index}] value={0} → {1}")
      @MethodSource
      @DisplayName("数値文字列を返す")
      void returnsExpected(double value, String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell)).isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(123.0, "123"),
            Arguments.of(123.45, "123.45"),
            Arguments.of(1.23456789012E11, "1.23457E11"));
      }
    }

    @Nested
    @DisplayName("NUMERIC セル（数値表示形式）のとき")
    class WhenCellTypeIsNumericWithNumberFormat {

      @ParameterizedTest(name = "[{index}] value={0}, format={1} → {2}")
      @MethodSource
      @DisplayName("数値書式で整形された文字列を返す")
      void returnsExpected(double value, String format, String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          CellStyle style = wb.createCellStyle();
          style.setDataFormat(wb.createDataFormat().getFormat(format));
          cell.setCellStyle(style);
          assertThat(ExcelReadUtil.getStringFromCell(cell)).isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(1234.5, "0.00", "1234.50"),
            Arguments.of(1234567.0, "#,##0", "1,234,567"),
            Arguments.of(0.1, "0%", "10%"));
      }
    }

    @Nested
    @DisplayName("NUMERIC セル（日付表示形式）のとき")
    class WhenCellTypeIsNumericWithDateFormat {

      @ParameterizedTest(name = "[{index}] dateTimeFormat={0} → {1}")
      @MethodSource
      @DisplayName("dateTimeFormat で整形された日付文字列を返す")
      void dateOnly(@Nullable DateTimeFormatter dateTimeFormat, String expected)
          throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(DateUtil.getExcelDate(LocalDate.of(2000, 1, 23)));
          CellStyle style = wb.createCellStyle();
          style.setDataFormat(wb.createDataFormat().getFormat("yyyy/mm/dd"));
          cell.setCellStyle(style);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, dateTimeFormat, null))
              .isEqualTo(expected);
        }
      }

      static Stream<Arguments> dateOnly() {
        return Stream.of(
            Arguments.of(null, "2000-01-23"),
            Arguments.of(DateTimeFormatter.ofPattern("yyyy/M/d"), "2000/1/23"));
      }

      @Test
      @DisplayName("日付＋時刻セルを dateTimeFormat で整形して返す")
      void dateTime() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(LocalDateTime.of(2000, 1, 23, 12, 34, 56));
          CellStyle style = wb.createCellStyle();
          style.setDataFormat(wb.createDataFormat().getFormat("yyyy/mm/dd hh:mm:ss"));
          cell.setCellStyle(style);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null,
              DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"), null))
              .isEqualTo("2000-01-23 12:34:56");
        }
      }
    }

    @Nested
    @DisplayName("ERROR セルのとき")
    class WhenCellTypeIsError {

      @Test
      @DisplayName("ExcelAppException がスローされる")
      void throwsExcelAppException() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet("Sheet1").createRow(0).createCell(0);
          cell.setCellErrorValue(FormulaError.NUM.getCode());
          assertThatThrownBy(() -> ExcelReadUtil.getStringFromCell(cell))
              .isInstanceOf(ExcelAppException.class);
        }
      }
    }

    @Nested
    @DisplayName("BOOLEAN セルのとき")
    class WhenCellTypeIsBoolean {

      @ParameterizedTest(name = "[{index}] value={0} → {1}")
      @MethodSource
      @DisplayName("\"TRUE\" または \"FALSE\" を返す")
      void returnsExpected(boolean value, String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell)).isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(true, "TRUE"),
            Arguments.of(false, "FALSE"));
      }
    }

    @Nested
    @DisplayName("FORMULA セルのとき")
    class WhenCellTypeIsFormula {

      @ParameterizedTest(name = "[{index}] formula={0} → {2}")
      @MethodSource
      @DisplayName("キャッシュ結果に応じた値を返す")
      void returnsExpected(String formula, @Nullable String noDataString,
          @Nullable String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellFormula(formula);
          wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, null, noDataString))
              .isEqualTo(expected);
        }
      }

      static Stream<Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of("\"hello\"", null, "hello"),
            Arguments.of("1+1", null, "2"),
            Arguments.of("\"\"", "N/A", "N/A"));
      }

      @Test
      @DisplayName("数式がエラー（#DIV/0! など）のとき ExcelAppException がスローされる")
      void whenFormulaReturnsError() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellFormula("1/0");
          wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
          assertThatThrownBy(() -> ExcelReadUtil.getStringFromCell(cell))
              .isInstanceOf(ExcelAppException.class);
        }
      }
    }
  }
}
