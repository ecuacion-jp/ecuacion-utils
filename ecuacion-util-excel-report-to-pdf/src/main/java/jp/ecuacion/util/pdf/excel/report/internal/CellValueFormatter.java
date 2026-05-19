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
package jp.ecuacion.util.pdf.excel.report.internal;

import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.FormatStyle;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.jspecify.annotations.Nullable;

/** Formats Excel cell values into display strings for PDF rendering. */
class CellValueFormatter {

  private static final LocalDate REIWA_START = LocalDate.of(2019, 5, 1);

  private final DataFormatter dataFormatter;
  private final Locale dateLocale;

  CellValueFormatter(DataFormatter dataFormatter, Locale dateLocale) {
    this.dataFormatter = dataFormatter;
    this.dateLocale = dateLocale;
  }

  /**
   * Returns the display value of a cell.
   *
   * <p>For formula cells, uses the cached result instead of the formula string,
   * since {@link DataFormatter#formatCellValue(Cell)} returns the formula string
   * when no {@code FormulaEvaluator} is provided.</p>
   *
   * <p>For numeric cells with date-like format strings (e.g. {@code yyyy"年"m"月分"}),
   * applies custom date formatting because POI's {@code DateUtil.isCellDateFormatted}
   * may return {@code false} for Japanese date formats.</p>
   */
  String getCellDisplayValue(Cell cell) {
    CellType effectiveType =
        (cell.getCellType() == CellType.FORMULA) ? cell.getCachedFormulaResultType()
            : cell.getCellType();

    if (effectiveType == CellType.NUMERIC) {
      String formatString = cell.getCellStyle().getDataFormatString();
      if (isLikelyDateFormat(formatString)) {
        double numVal = cell.getNumericCellValue();
        String builtin = formatBuiltinDateValue(cell.getCellStyle().getDataFormat(), numVal);
        if (builtin != null) {
          return builtin;
        }
        return formatDateValue(numVal, formatString);
      }
      if (cell.getCellType() == CellType.FORMULA) {
        double numericValue = cell.getNumericCellValue();
        // For zero values with a multi-section format, DataFormatter renders the "??" digit
        // placeholders in the zero section (e.g. `"-"??`) as actual digits ("- 0") instead
        // of alignment-only spaces. Extract just the quoted literals from the zero section.
        if (numericValue == 0.0) {
          String zeroLiteral = extractZeroSectionLiteral(formatString);
          if (zeroLiteral != null) {
            return zeroLiteral;
          }
        }
        return dataFormatter.formatRawCellContents(numericValue,
            cell.getCellStyle().getDataFormat(), formatString);
      }
      return dataFormatter.formatCellValue(cell);
    }

    if (cell.getCellType() == CellType.FORMULA) {
      return switch (cell.getCachedFormulaResultType()) {
        case STRING -> cell.getRichStringCellValue().getString();
        case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
        default -> "";
      };
    }

    return dataFormatter.formatCellValue(cell);
  }

  /**
   * Returns the quoted-literal content from the zero section of a multi-section Excel format
   * string, or {@code null} if no suitable zero section is found.
   *
   * <p>Excel accounting formats use a three- or four-section pattern such as
   * {@code _ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ } where the third section (for
   * zero) contains a quoted dash literal followed by {@code ??} digit-alignment placeholders.
   * {@link DataFormatter#formatRawCellContents} incorrectly renders those {@code ?} tokens as
   * the digit {@code 0}, producing {@code "- 0"} instead of the correct {@code "-"}.
   * This method extracts only the quoted-literal characters, skipping all format tokens.</p>
   */
  @Nullable String extractZeroSectionLiteral(String formatString) {
    List<String> sections = new ArrayList<>();
    int start = 0;
    boolean inQuote = false;
    for (int i = 0; i < formatString.length(); i++) {
      char c = formatString.charAt(i);
      if (c == '"') {
        inQuote = !inQuote;
      } else if (c == ';' && !inQuote) {
        sections.add(formatString.substring(start, i));
        start = i + 1;
      }
    }
    sections.add(formatString.substring(start));

    if (sections.size() < 3) {
      return null;
    }

    String zeroSection = sections.get(2);
    StringBuilder literal = new StringBuilder();
    boolean inQ = false;
    for (int i = 0; i < zeroSection.length(); i++) {
      char c = zeroSection.charAt(i);
      if (c == '"') {
        inQ = !inQ;
      } else if (inQ) {
        literal.append(c);
      } else if (c == '\\' && i + 1 < zeroSection.length()) {
        literal.append(zeroSection.charAt(++i));
      }
    }

    return literal.length() > 0 ? literal.toString() : null;
  }

  /**
   * Formats an Excel date serial number using JRE locale-aware formatting for built-in
   * date format IDs.
   *
   * <p>Excel's built-in date formats (e.g. ID 14 = {@code m/d/yy}) are locale-sensitive:
   * Excel renders them according to the OS locale rather than the stored format string.
   * POI always returns the en-US format string, so we use
   * {@link DateTimeFormatter#ofLocalizedDate} with {@link #dateLocale} to reproduce
   * Excel's behaviour for any locale without maintaining a hand-coded mapping table.</p>
   *
   * @return locale-formatted date string, or {@code null} if this ID is not handled here
   */
  @Nullable String formatBuiltinDateValue(short formatId, double numericValue) {
    FormatStyle style = switch (formatId) {
      case 14 -> FormatStyle.SHORT;
      default -> null;
    };
    if (style == null) {
      return null;
    }
    LocalDate date = DateUtil.getLocalDateTime(numericValue, false).toLocalDate();
    return DateTimeFormatter.ofLocalizedDate(style).withLocale(dateLocale).format(date);
  }

  /**
   * Returns {@code true} if the format string looks like a date format.
   *
   * <p>POI's {@code DateUtil.isCellDateFormatted} can return {@code false} for
   * Japanese date format strings such as {@code yyyy"年"m"月分"}.
   * This method checks for year tokens ({@code y}/{@code Y}) outside of quoted
   * literals as a more lenient detection heuristic.</p>
   */
  boolean isLikelyDateFormat(String formatString) {
    if (formatString == null || formatString.isEmpty()
        || formatString.equalsIgnoreCase("General")) {
      return false;
    }
    String mainSection =
        formatString.contains(";") ? formatString.substring(0, formatString.indexOf(';'))
            : formatString;
    String stripped = mainSection.replaceAll("\"[^\"]*\"", "");
    stripped = stripped.replaceAll("\\[[^\\]]*\\]", "");
    return stripped.contains("y") || stripped.contains("Y") || stripped.contains("g")
        || stripped.contains("G");
  }

  /** Formats an Excel date serial number using the given Excel format string. */
  String formatDateValue(double numericValue, String formatString) {
    var ldt = DateUtil.getLocalDateTime(numericValue, false);
    LocalDate date = ldt.toLocalDate();
    LocalTime time = ldt.toLocalTime();
    Locale locale = extractFormatLocale(formatString);

    String fmt = formatString.contains(";") ? formatString.substring(0, formatString.indexOf(';'))
        : formatString;

    boolean is12Hour = fmt.toLowerCase(Locale.ENGLISH).contains("am/pm")
        || fmt.toLowerCase(Locale.ENGLISH).contains("a/p");
    int hourVal = is12Hour ? (time.getHour() % 12 == 0 ? 12 : time.getHour() % 12) : time.getHour();

    StringBuilder result = new StringBuilder();
    boolean lastWasHour = false;
    int i = 0;

    while (i < fmt.length()) {
      char c = fmt.charAt(i);
      if (c == '"') {
        int end = fmt.indexOf('"', i + 1);
        if (end > i) {
          result.append(fmt, i + 1, end);
          i = end + 1;
        } else {
          i++;
        }
      } else if (c == '[') {
        int end = fmt.indexOf(']', i);
        i = (end > i) ? end + 1 : i + 1;
      } else if (c == 'y' || c == 'Y') {
        int n = countConsecutive(fmt, i, c);
        result.append(n >= 4 ? String.format("%04d", date.getYear())
            : String.format("%02d", date.getYear() % 100));
        i += n;
        lastWasHour = false;
      } else if (c == 'g' || c == 'G') {
        int n = countConsecutive(fmt, i, c);
        result.append(eraName(date, n));
        i += n;
        lastWasHour = false;
      } else if (c == 'e' || c == 'E') {
        int n = countConsecutive(fmt, i, c);
        result.append(eraYear(date, n));
        i += n;
        lastWasHour = false;
      } else if (c == 'm' || c == 'M') {
        int n = countConsecutive(fmt, i, c);
        if (lastWasHour) {
          result.append(
              n >= 2 ? String.format("%02d", time.getMinute()) : String.valueOf(time.getMinute()));
          lastWasHour = false;
        } else if (n >= 4) {
          result.append(date.getMonth().getDisplayName(TextStyle.FULL, locale));
        } else if (n == 3) {
          result.append(date.getMonth().getDisplayName(TextStyle.SHORT, locale));
        } else {
          result.append(n >= 2 ? String.format("%02d", date.getMonthValue())
              : String.valueOf(date.getMonthValue()));
        }
        i += n;
      } else if (c == 'd' || c == 'D') {
        int n = countConsecutive(fmt, i, c);
        if (n >= 4) {
          result.append(date.getDayOfWeek().getDisplayName(TextStyle.FULL, locale));
        } else if (n == 3) {
          result.append(date.getDayOfWeek().getDisplayName(TextStyle.SHORT, locale));
        } else {
          result.append(n >= 2 ? String.format("%02d", date.getDayOfMonth())
              : String.valueOf(date.getDayOfMonth()));
        }
        i += n;
        lastWasHour = false;
      } else if (c == 'h' || c == 'H') {
        int n = countConsecutive(fmt, i, c);
        result.append(n >= 2 ? String.format("%02d", hourVal) : String.valueOf(hourVal));
        i += n;
        lastWasHour = true;
      } else if (c == 's' || c == 'S') {
        int n = countConsecutive(fmt, i, c);
        result.append(
            n >= 2 ? String.format("%02d", time.getSecond()) : String.valueOf(time.getSecond()));
        i += n;
        lastWasHour = false;
      } else if (c == 'a' || c == 'A') {
        if (i + 4 < fmt.length() && fmt.substring(i, i + 5).equalsIgnoreCase("AM/PM")) {
          result.append(time.getHour() < 12 ? "AM" : "PM");
          i += 5;
        } else if (i + 2 < fmt.length() && fmt.substring(i, i + 3).equalsIgnoreCase("A/P")) {
          result.append(time.getHour() < 12 ? "A" : "P");
          i += 3;
        } else {
          int n = countConsecutive(fmt, i, c);
          if (n >= 3) {
            int dow = date.getDayOfWeek().getValue() - 1;
            String[] s = {"月", "火", "水", "木", "金", "土", "日"};
            String[] l = {"月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"};
            result.append(n >= 4 ? l[dow] : s[dow]);
            i += n;
          } else {
            result.append(c);
            i++;
          }
        }
      } else {
        result.append(c);
        i++;
      }
    }
    return result.toString();
  }

  private Locale extractFormatLocale(String formatString) {
    int start = formatString.indexOf("[$-");
    if (start >= 0) {
      int end = formatString.indexOf(']', start + 3);
      if (end > start + 3) {
        try {
          int lcid = Integer.parseInt(formatString.substring(start + 3, end), 16) & 0xFFFF;
          if (lcid == 0x0411) return Locale.JAPANESE;
          if (lcid == 0x0407) return Locale.GERMAN;
          if (lcid == 0x040C) return Locale.FRENCH;
          return Locale.ENGLISH;
        } catch (NumberFormatException ignored) {
          // fall through
        }
      }
    }
    return Locale.ENGLISH;
  }

  private String eraName(LocalDate date, int count) {
    if (date.isBefore(REIWA_START)) {
      throw new RuntimeException(
          "Japanese era before Reiwa (2019-05-01) is not supported: " + date);
    }
    return count >= 2 ? "令和" : "令";
  }

  private String eraYear(LocalDate date, int count) {
    if (date.isBefore(REIWA_START)) {
      throw new RuntimeException(
          "Japanese era before Reiwa (2019-05-01) is not supported: " + date);
    }
    int year = date.getYear() - 2018;
    return count >= 2 ? String.format("%02d", year) : String.valueOf(year);
  }

  private int countConsecutive(String s, int start, char target) {
    char lower = Character.toLowerCase(target);
    int count = 0;
    while (start + count < s.length() && Character.toLowerCase(s.charAt(start + count)) == lower) {
      count++;
    }
    return count;
  }
}
