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

/** Provides character classification for line breaking and spacing of East Asian text. */
final class TextRules {

  /** Characters which must not start a line (Word's standard kinsoku setting for Japanese). */
  private static final String LINE_START_PROHIBITED =
      "!%),.:;?]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠";
  /** Characters which must not end a line. */
  private static final String LINE_END_PROHIBITED = "$([\\{£¥‘“〈《「『【〔＄（［｛｢￡￥";
  /** Punctuation which may hang beyond the right edge of a line. */
  private static final String HANGABLE = "、。，．";
  private static final String FULL_WIDTH_OPENING = "（〔［｛〈《「『【‘“";
  private static final String FULL_WIDTH_CLOSING = "、。，．）〕］｝〉》」』】’”";

  private TextRules() {}

  static boolean lineStartProhibited(int cp) {
    return cp < 0x10000 && LINE_START_PROHIBITED.indexOf(cp) >= 0;
  }

  static boolean lineEndProhibited(int cp) {
    return cp < 0x10000 && LINE_END_PROHIBITED.indexOf(cp) >= 0;
  }

  static boolean hangable(int cp) {
    return cp < 0x10000 && HANGABLE.indexOf(cp) >= 0;
  }

  static boolean fullWidthOpening(int cp) {
    return cp >= 0x2000 && cp < 0x10000 && FULL_WIDTH_OPENING.indexOf(cp) >= 0;
  }

  static boolean fullWidthClosing(int cp) {
    return cp >= 0x2000 && cp < 0x10000 && FULL_WIDTH_CLOSING.indexOf(cp) >= 0;
  }

  /** Returns whether the character is East Asian, i.e. a line can break next to it. */
  static boolean eastAsian(int cp) {
    return (cp >= 0x1100 && cp <= 0x11FF) || (cp >= 0x2E80 && cp <= 0xA4CF)
        || (cp >= 0xAC00 && cp <= 0xD7FF) || (cp >= 0xF900 && cp <= 0xFAFF)
        || (cp >= 0xFE30 && cp <= 0xFE4F) || (cp >= 0xFF00 && cp <= 0xFFEF) || cp >= 0x20000;
  }

  /** Returns whether the character is an East Asian letter (ideograph, kana, hangul). */
  static boolean eastAsianLetter(int cp) {
    return (cp >= 0x3040 && cp <= 0x30FF && cp != 0x30FB && cp != 0x30FC)
        || (cp >= 0x3400 && cp <= 0x9FFF)
        || (cp >= 0xAC00 && cp <= 0xD7FF) || (cp >= 0xF900 && cp <= 0xFAFF) || cp >= 0x20000
        || cp == 0x3005;
  }

  static boolean latinLetter(int cp) {
    return cp < 0x2000 && Character.isLetter(cp);
  }

  static boolean digit(int cp) {
    return cp >= '0' && cp <= '9';
  }

  /** Returns whether a line can break between the two adjacent characters. */
  static boolean canBreak(int before, int after, boolean wordWrap) {
    if (after == ' ' || after == 0x00A0 || before == 0x00A0 || after == 0x2011
        || before == 0x2011) {
      return false;
    }
    if (before == ' ' || before == 0x3000) {
      return true;
    }
    if (lineStartProhibited(after) || lineEndProhibited(before)) {
      return false;
    }
    if (eastAsian(before) || eastAsian(after)) {
      return true;
    }
    if (before == '-' && Character.isLetter(after)) {
      return true;
    }
    return !wordWrap;
  }
}
