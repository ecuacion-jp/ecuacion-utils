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
package jp.ecuacion.util.excel.table.bean;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import org.jspecify.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

@DisplayName("StringExcelTableBean")
public class StringExcelTableBeanTest {

  private static List<String> nullList() {
    List<String> list = new ArrayList<>();
    list.add(null);
    return list;
  }

  static class AllTypesBean extends StringExcelTableBean {
    @Nullable Short shortField;
    @Nullable Float floatField;
    @Nullable Double doubleField;
    @Nullable BigDecimal bigDecimalField;
    @Nullable BigInteger bigIntegerField;
    @Nullable Boolean booleanField;
    @Nullable LocalDate localDateField;
    @Nullable LocalDateTime localDateTimeField;
    @Nullable LocalTime localTimeField;

    public AllTypesBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"shortField", "floatField", "doubleField", "bigDecimalField",
          "bigIntegerField", "booleanField", "localDateField", "localDateTimeField",
          "localTimeField"};
    }
  }

  static class ShortBean extends StringExcelTableBean {
    @Nullable Short value;

    public ShortBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class FloatBean extends StringExcelTableBean {
    @Nullable Float value;

    public FloatBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class DoubleBean extends StringExcelTableBean {
    @Nullable Double value;

    public DoubleBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class BigDecimalBean extends StringExcelTableBean {
    @Nullable BigDecimal value;

    public BigDecimalBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class BigIntegerBean extends StringExcelTableBean {
    @Nullable BigInteger value;

    public BigIntegerBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class BooleanBean extends StringExcelTableBean {
    @Nullable Boolean value;

    public BooleanBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class LocalDateBean extends StringExcelTableBean {
    @Nullable LocalDate value;

    public LocalDateBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class LocalDateTimeBean extends StringExcelTableBean {
    @Nullable LocalDateTime value;

    public LocalDateTimeBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class LocalTimeBean extends StringExcelTableBean {
    @Nullable LocalTime value;

    public LocalTimeBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }
  }

  static class CustomDateFormatBean extends StringExcelTableBean {
    @Nullable LocalDate value;

    public CustomDateFormatBean(List<String> colList) {
      super(colList);
    }

    @Override
    protected @Nullable String[] getFieldNameArray() {
      return new String[] {"value"};
    }

    @Override
    protected DateTimeFormatter getDateTimeFormatter() {
      return DateTimeFormatter.ofPattern("yyyy/MM/dd");
    }
  }

  @Nested
  @DisplayName("convertToFieldType: numeric types")
  class NumericTypes {

    @Test
    @DisplayName("Short field → parsed correctly")
    void shortField() {
      var bean = new ShortBean(List.of("42"));
      assertThat(bean.value).isEqualTo((short) 42);
    }

    @Test
    @DisplayName("Short field: null input → null")
    void shortFieldNull() {
      var bean = new ShortBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("Float field → parsed correctly")
    void floatField() {
      var bean = new FloatBean(List.of("3.14"));
      assertThat(bean.value).isCloseTo(3.14f, org.assertj.core.data.Offset.offset(0.001f));
    }

    @Test
    @DisplayName("Float field: null input → null")
    void floatFieldNull() {
      var bean = new FloatBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("Double field → parsed correctly")
    void doubleField() {
      var bean = new DoubleBean(List.of("2.718"));
      assertThat(bean.value).isCloseTo(2.718, org.assertj.core.data.Offset.offset(0.0001));
    }

    @Test
    @DisplayName("Double field: null input → null")
    void doubleFieldNull() {
      var bean = new DoubleBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("BigDecimal field → parsed correctly")
    void bigDecimalField() {
      var bean = new BigDecimalBean(List.of("123.456"));
      assertThat(bean.value).isEqualByComparingTo(new BigDecimal("123.456"));
    }

    @Test
    @DisplayName("BigDecimal field: null input → null")
    void bigDecimalFieldNull() {
      var bean = new BigDecimalBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("BigInteger field → parsed correctly")
    void bigIntegerField() {
      var bean = new BigIntegerBean(List.of("999999999999"));
      assertThat(bean.value).isEqualTo(new BigInteger("999999999999"));
    }

    @Test
    @DisplayName("BigInteger field: null input → null")
    void bigIntegerFieldNull() {
      var bean = new BigIntegerBean(nullList());
      assertThat(bean.value).isNull();
    }
  }

  @Nested
  @DisplayName("convertToFieldType: Boolean")
  class BooleanType {

    @Test
    @DisplayName("\"true\" → Boolean.TRUE")
    void trueValue() {
      var bean = new BooleanBean(List.of("true"));
      assertThat(bean.value).isTrue();
    }

    @Test
    @DisplayName("\"false\" → Boolean.FALSE")
    void falseValue() {
      var bean = new BooleanBean(List.of("false"));
      assertThat(bean.value).isFalse();
    }

    @Test
    @DisplayName("null input → null")
    void nullValue() {
      var bean = new BooleanBean(nullList());
      assertThat(bean.value).isNull();
    }
  }

  @Nested
  @DisplayName("convertToFieldType: date/time types")
  class DateTimeTypes {

    @Test
    @DisplayName("LocalDate field → parsed with ISO format")
    void localDateField() {
      var bean = new LocalDateBean(List.of("2024-03-15"));
      assertThat(bean.value).isEqualTo(LocalDate.of(2024, 3, 15));
    }

    @Test
    @DisplayName("LocalDate field: null input → null")
    void localDateFieldNull() {
      var bean = new LocalDateBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("LocalDate field: custom DateTimeFormatter via override")
    void localDateCustomFormat() {
      var bean = new CustomDateFormatBean(List.of("2024/03/15"));
      assertThat(bean.value).isEqualTo(LocalDate.of(2024, 3, 15));
    }

    @Test
    @DisplayName("LocalDateTime field → parsed with ISO format")
    void localDateTimeField() {
      var bean = new LocalDateTimeBean(List.of("2024-03-15T10:30:00"));
      assertThat(bean.value).isEqualTo(LocalDateTime.of(2024, 3, 15, 10, 30, 0));
    }

    @Test
    @DisplayName("LocalDateTime field: null input → null")
    void localDateTimeFieldNull() {
      var bean = new LocalDateTimeBean(nullList());
      assertThat(bean.value).isNull();
    }

    @Test
    @DisplayName("LocalTime field → parsed correctly")
    void localTimeField() {
      var bean = new LocalTimeBean(List.of("10:30:00"));
      assertThat(bean.value).isEqualTo(LocalTime.of(10, 30, 0));
    }

    @Test
    @DisplayName("LocalTime field: null input → null")
    void localTimeFieldNull() {
      var bean = new LocalTimeBean(nullList());
      assertThat(bean.value).isNull();
    }
  }

  @Nested
  @DisplayName("convertToFieldType: conversion errors")
  class ConversionErrors {

    @Test
    @DisplayName("non-numeric string for Integer field → RuntimeException with field name")
    void invalidInteger() {
      assertThatThrownBy(() -> new ShortBean(List.of("not-a-number")))
          .isInstanceOf(RuntimeException.class)
          .hasMessageContaining("value");
    }

    @Test
    @DisplayName("invalid date string for LocalDate → RuntimeException with field name")
    void invalidLocalDate() {
      assertThatThrownBy(() -> new LocalDateBean(List.of("not-a-date")))
          .isInstanceOf(RuntimeException.class)
          .hasMessageContaining("value");
    }
  }
}
