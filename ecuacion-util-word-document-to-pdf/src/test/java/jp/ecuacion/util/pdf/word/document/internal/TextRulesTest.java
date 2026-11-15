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

@DisplayName("TextRules")
class TextRulesTest {

  @ParameterizedTest(name = "[{0}][{1}] -> {2}")
  @CsvSource({
      // East Asian characters: a line can break anywhere between them
      "漢, 字, true", "あ, ア, true", "漢, a, true", "a, 漢, true", "1, 円, true",
      // Latin words and numbers are not divided
      "a, b, false", "1, 2, false", "a, 1, false", "1, ',', false",
      // a space belongs to the word before it
      "a, ' ', false", "' ', a, true",
      // after a hyphen
      "-, a, true", "-, 1, false",
      // characters which must not start a line
      "字, 。, false", "字, 、, false", "字, ）, false", "字, 」, false", "a, ), false",
      "字, ？, false", "字, ・, false",
      // characters which must not end a line
      "（, 字, false", "「, 字, false", "(, a, false",
      // after closing punctuation, before opening punctuation
      "。, 字, true", "）, 字, true", "字, （, true", "」, 「, true"})
  @DisplayName("canBreak")
  void canBreak(String before, String after, boolean expected) {
    assertThat(TextRules.canBreak(before.codePointAt(0), after.codePointAt(0), true))
        .isEqualTo(expected);
  }

  @ParameterizedTest(name = "wordWrap off: [{0}][{1}] -> {2}")
  @CsvSource({"a, b, true", "字, 。, false", "（, 字, false"})
  @DisplayName("canBreak without word wrap: Latin words can be divided, kinsoku still applies")
  void canBreakWithoutWordWrap(String before, String after, boolean expected) {
    assertThat(TextRules.canBreak(before.codePointAt(0), after.codePointAt(0), false))
        .isEqualTo(expected);
  }
}
