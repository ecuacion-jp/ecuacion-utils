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
package jp.ecuacion.util.pdf.word.document.internal;

import static org.assertj.core.api.Assertions.assertThat;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

@DisplayName("NumberFormats")
class NumberFormatsTest {

  @ParameterizedTest(name = "{0}: {1} -> {2}")
  @CsvSource({"decimal, 1, 1", "decimal, 12, 12", "decimalZero, 3, 03", "decimalZero, 12, 12",
      "upperRoman, 4, IV", "upperRoman, 1994, MCMXCIV", "lowerRoman, 9, ix",
      "upperLetter, 1, A", "upperLetter, 26, Z", "upperLetter, 27, AA", "lowerLetter, 28, bb",
      "decimalFullWidth, 10, １０", "decimalEnclosedCircle, 1, ①", "decimalEnclosedCircle, 20, ⑳",
      "decimalEnclosedCircle, 21, 21", "decimalEnclosedParen, 2, ⑵",
      "japaneseCounting, 1, 一", "japaneseCounting, 10, 十", "japaneseCounting, 11, 十一",
      "japaneseCounting, 20, 二十", "japaneseCounting, 123, 百二十三",
      "japaneseCounting, 2025, 二千二十五", "japaneseLegal, 12, 拾弐",
      "ideographDigital, 10, 一〇", "aiueoFullWidth, 1, ア", "aiueoFullWidth, 5, オ",
      "irohaFullWidth, 3, ハ", "aiueo, 2, ｲ", "iroha, 1, ｲ", "none, 5, ''",
      "unknownFormat, 7, 7"})
  @DisplayName("format")
  void format(String format, int number, String expected) {
    assertThat(NumberFormats.format(number, format)).isEqualTo(expected);
  }
}
