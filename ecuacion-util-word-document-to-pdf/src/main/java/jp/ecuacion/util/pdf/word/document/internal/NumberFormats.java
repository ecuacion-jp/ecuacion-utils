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

/** Formats numbers in the number formats of WordprocessingML ({@code w:numFmt}). */
final class NumberFormats {

  private static final String KANJI_DIGITS = "〇一二三四五六七八九";
  private static final String LEGAL_DIGITS = "〇壱弐参四五六七八九";
  private static final String AIUEO =
      "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン";
  private static final String IROHA =
      "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス";
  private static final String AIUEO_HALF =
      "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ";
  private static final String IROHA_HALF =
      "ｲﾛﾊﾆﾎﾍﾄﾁﾘﾇﾙｦﾜｶﾖﾀﾚｿﾂﾈﾅﾗﾑｳｲﾉｵｸﾔﾏｹﾌｺｴﾃｱｻｷﾕﾒﾐｼｴﾋﾓｾｽ";

  private NumberFormats() {}

  static String format(int n, String format) {
    return switch (format) {
      case "none" -> "";
      case "decimalZero" -> n < 10 && n >= 0 ? "0" + n : Integer.toString(n);
      case "upperRoman" -> roman(n);
      case "lowerRoman" -> roman(n).toLowerCase(java.util.Locale.ROOT);
      case "upperLetter" -> letters(n, 'A');
      case "lowerLetter" -> letters(n, 'a');
      case "decimalFullWidth", "decimalFullWidth2" -> fullWidth(n);
      case "decimalEnclosedCircle", "decimalEnclosedCircleChinese" ->
          n >= 1 && n <= 20 ? String.valueOf((char) (0x2460 + n - 1)) : Integer.toString(n);
      case "decimalEnclosedParen" ->
          n >= 1 && n <= 20 ? String.valueOf((char) (0x2474 + n - 1)) : Integer.toString(n);
      case "decimalEnclosedFullstop" ->
          n >= 1 && n <= 20 ? String.valueOf((char) (0x2488 + n - 1)) : Integer.toString(n);
      case "japaneseCounting", "chineseCounting", "chineseCountingThousand",
          "ideographTraditional", "taiwaneseCounting", "taiwaneseCountingThousand" ->
          kanji(n, KANJI_DIGITS, "十", "百", "千");
      case "japaneseLegal", "chineseLegalSimplified" -> kanji(n, LEGAL_DIGITS, "拾", "百", "阡");
      case "japaneseDigitalTenThousand", "ideographDigital", "taiwaneseDigital" ->
          digits(n, KANJI_DIGITS);
      case "aiueoFullWidth" -> cycle(n, AIUEO);
      case "irohaFullWidth" -> cycle(n, IROHA);
      case "aiueo" -> cycle(n, AIUEO_HALF);
      case "iroha" -> cycle(n, IROHA_HALF);
      default -> Integer.toString(n);
    };
  }

  private static String roman(int n) {
    if (n <= 0 || n >= 4000) {
      return Integer.toString(n);
    }
    int[] values = {1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1};
    String[] symbols = {"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"};
    StringBuilder sb = new StringBuilder();
    int rest = n;
    for (int i = 0; i < values.length; i++) {
      while (rest >= values[i]) {
        sb.append(symbols[i]);
        rest -= values[i];
      }
    }
    return sb.toString();
  }

  /** A, B, ..., Z, AA, BB, ... (Word repeats the letter). */
  private static String letters(int n, char base) {
    if (n <= 0) {
      return Integer.toString(n);
    }
    char c = (char) (base + (n - 1) % 26);
    return String.valueOf(c).repeat((n - 1) / 26 + 1);
  }

  private static String fullWidth(int n) {
    StringBuilder sb = new StringBuilder();
    for (char c : Integer.toString(n).toCharArray()) {
      sb.append(c >= '0' && c <= '9' ? (char) ('０' + (c - '0')) : c);
    }
    return sb.toString();
  }

  private static String digits(int n, String digitChars) {
    StringBuilder sb = new StringBuilder();
    for (char c : Integer.toString(n).toCharArray()) {
      sb.append(c >= '0' && c <= '9' ? digitChars.charAt(c - '0') : c);
    }
    return sb.toString();
  }

  private static String kanji(int n, String digitChars, String ten, String hundred,
      String thousand) {
    if (n <= 0 || n >= 10000) {
      return n == 0 ? digitChars.substring(0, 1) : Integer.toString(n);
    }
    StringBuilder sb = new StringBuilder();
    int[] units = {1000, 100, 10};
    String[] names = {thousand, hundred, ten};
    int rest = n;
    for (int i = 0; i < units.length; i++) {
      int d = rest / units[i];
      rest %= units[i];
      if (d > 0) {
        if (d > 1) {
          sb.append(digitChars.charAt(d));
        }
        sb.append(names[i]);
      }
    }
    if (rest > 0) {
      sb.append(digitChars.charAt(rest));
    }
    return sb.toString();
  }

  private static String cycle(int n, String chars) {
    if (n <= 0) {
      return Integer.toString(n);
    }
    return String.valueOf(chars.charAt((n - 1) % chars.length()));
  }
}
