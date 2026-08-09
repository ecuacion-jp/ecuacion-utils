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
import jp.ecuacion.util.excel.exception.ExcelTableException;
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
    @DisplayName("when value is null or an empty string")
    class WhenValueIsNullOrEmpty {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("returns noDataString as is")
      void returnsNoDataString(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) {
        assertThat(ExcelReadUtil.getNoDataStringIfNoData(value, noDataString))
            .isEqualTo(expected);
      }

      static @Nullable Stream<@Nullable Arguments> returnsNoDataString() {
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
    @DisplayName("when value is a normal string")
    class WhenValueIsNotEmpty {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("returns value as is (noDataString is ignored)")
      void returnsValue(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) {
        assertThat(ExcelReadUtil.getNoDataStringIfNoData(value, noDataString))
            .isEqualTo(expected);
      }

      static @Nullable Stream<@Nullable Arguments> returnsValue() {
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
    @DisplayName("when cell is null")
    class WhenCellIsNull {

      @ParameterizedTest(name = "[{index}] noDataString={0} → {1}")
      @MethodSource
      @DisplayName("returns noDataString")
      void returnsNoDataString(@Nullable String noDataString, @Nullable String expected)
          throws ExcelTableException {
        assertThat(ExcelReadUtil.getStringFromCell(null, null, null, noDataString))
            .isEqualTo(expected);
      }

      static @Nullable Stream<@Nullable Arguments> returnsNoDataString() {
        return Stream.of(
            Arguments.of(null, null),
            Arguments.of("N/A", "N/A"));
      }
    }

    @Nested
    @DisplayName("when cell type is BLANK")
    class WhenCellTypeIsBlank {

      @ParameterizedTest(name = "[{index}] noDataString={0} → {1}")
      @MethodSource
      @DisplayName("returns noDataString")
      void returnsNoDataString(@Nullable String noDataString, @Nullable String expected)
          throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, null, noDataString))
              .isEqualTo(expected);
        }
      }

      static @Nullable Stream<@Nullable Arguments> returnsNoDataString() {
        return Stream.of(
            Arguments.of(null, null),
            Arguments.of("N/A", "N/A"));
      }
    }

    @Nested
    @DisplayName("when cell type is STRING")
    class WhenCellTypeIsString {

      @ParameterizedTest(name = "[{index}] value={0}, noDataString={1} → {2}")
      @MethodSource
      @DisplayName("returns the string value or noDataString")
      void returnsExpected(@Nullable String value, @Nullable String noDataString,
          @Nullable String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell, null, null, noDataString))
              .isEqualTo(expected);
        }
      }

      static @Nullable Stream<@Nullable Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of("hello", null, "hello"),
            Arguments.of("", null, null),
            Arguments.of("", "N/A", "N/A"),
            Arguments.of(" ", null, " "));
      }
    }

    @Nested
    @DisplayName("when cell type is NUMERIC (format: General)")
    class WhenCellTypeIsNumericWithGeneralFormat {

      @ParameterizedTest(name = "[{index}] value={0} → {1}")
      @MethodSource
      @DisplayName("returns the numeric string")
      void returnsExpected(double value, String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell)).isEqualTo(expected);
        }
      }

      static @Nullable Stream<@Nullable Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(123.0, "123"),
            Arguments.of(123.45, "123.45"),
            Arguments.of(1.23456789012E11, "1.23457E11"));
      }
    }

    @Nested
    @DisplayName("when cell type is NUMERIC (number format)")
    class WhenCellTypeIsNumericWithNumberFormat {

      @ParameterizedTest(name = "[{index}] value={0}, format={1} → {2}")
      @MethodSource
      @DisplayName("returns the string formatted with the number format")
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

      static @Nullable Stream<@Nullable Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(1234.5, "0.00", "1234.50"),
            Arguments.of(1234567.0, "#,##0", "1,234,567"),
            Arguments.of(0.1, "0%", "10%"));
      }
    }

    @Nested
    @DisplayName("when cell type is NUMERIC (date format)")
    class WhenCellTypeIsNumericWithDateFormat {

      @ParameterizedTest(name = "[{index}] dateTimeFormat={0} → {1}")
      @MethodSource
      @DisplayName("returns the date string formatted with dateTimeFormat")
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

      static @Nullable Stream<@Nullable Arguments> dateOnly() {
        return Stream.of(
            Arguments.of(null, "2000-01-23"),
            Arguments.of(DateTimeFormatter.ofPattern("yyyy/M/d"), "2000/1/23"));
      }

      @Test
      @DisplayName("formats a date+time cell with dateTimeFormat and returns it")
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
    @DisplayName("when cell type is ERROR")
    class WhenCellTypeIsError {

      @Test
      @DisplayName("throws ExcelTableException")
      void throwsExcelTableException() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet("Sheet1").createRow(0).createCell(0);
          cell.setCellErrorValue(FormulaError.NUM.getCode());
          assertThatThrownBy(() -> ExcelReadUtil.getStringFromCell(cell))
              .isInstanceOf(ExcelTableException.class);
        }
      }
    }

    @Nested
    @DisplayName("when cell type is BOOLEAN")
    class WhenCellTypeIsBoolean {

      @ParameterizedTest(name = "[{index}] value={0} → {1}")
      @MethodSource
      @DisplayName("returns \"TRUE\" or \"FALSE\"")
      void returnsExpected(boolean value, String expected) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellValue(value);
          assertThat(ExcelReadUtil.getStringFromCell(cell)).isEqualTo(expected);
        }
      }

      static @Nullable Stream<@Nullable Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of(true, "TRUE"),
            Arguments.of(false, "FALSE"));
      }
    }

    @Nested
    @DisplayName("when cell type is FORMULA")
    class WhenCellTypeIsFormula {

      @ParameterizedTest(name = "[{index}] formula={0} → {2}")
      @MethodSource
      @DisplayName("returns the value based on the cached result")
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

      static @Nullable Stream<@Nullable Arguments> returnsExpected() {
        return Stream.of(
            Arguments.of("\"hello\"", null, "hello"),
            Arguments.of("1+1", null, "2"),
            Arguments.of("\"\"", "N/A", "N/A"));
      }

      @Test
      @DisplayName("throws ExcelTableException when the formula is an error (e.g. #DIV/0!)")
      void whenFormulaReturnsError() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
          Cell cell = wb.createSheet().createRow(0).createCell(0);
          cell.setCellFormula("1/0");
          wb.getCreationHelper().createFormulaEvaluator().evaluateFormulaCell(cell);
          assertThatThrownBy(() -> ExcelReadUtil.getStringFromCell(cell))
              .isInstanceOf(ExcelTableException.class);
        }
      }
    }
  }
}
