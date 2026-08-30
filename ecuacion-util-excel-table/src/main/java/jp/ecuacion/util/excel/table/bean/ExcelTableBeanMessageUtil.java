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

import org.jspecify.annotations.Nullable;

/**
 * Formats Excel cell values for safe inclusion in exception messages.
 */
final class ExcelTableBeanMessageUtil {

  private static final int MAX_VALUE_LENGTH_IN_MESSAGE = 200;

  private ExcelTableBeanMessageUtil() {}

  /**
   * Returns a truncated, control-character-escaped representation of {@code value}, safe to
   * embed in an exception message even when {@code value} originates from an untrusted Excel
   * cell.
   *
   * <p><strong>Security note:</strong> without this, a crafted cell value containing newlines
   *     could forge fake log lines (log injection) once the caller logs the resulting exception
   *     message.</p>
   *
   * @param value the value to format, may be {@code null}
   * @return a safe string representation
   */
  static String toSafeMessagePart(@Nullable Object value) {
    if (value == null) {
      return "null";
    }

    String s = value.toString();
    int len = Math.min(s.length(), MAX_VALUE_LENGTH_IN_MESSAGE);
    StringBuilder sb = new StringBuilder(len + 16);
    for (int i = 0; i < len; i++) {
      char c = s.charAt(i);
      switch (c) {
        case '\r' -> sb.append("\\r");
        case '\n' -> sb.append("\\n");
        case '\t' -> sb.append("\\t");
        default -> {
          if (c < 0x20) {
            sb.append(String.format("\\u%04x", (int) c));
          } else {
            sb.append(c);
          }
        }
      }
    }
    if (s.length() > MAX_VALUE_LENGTH_IN_MESSAGE) {
      sb.append("...(truncated)");
    }
    return sb.toString();
  }
}
