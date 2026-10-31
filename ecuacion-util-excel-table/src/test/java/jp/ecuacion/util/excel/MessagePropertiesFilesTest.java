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
package jp.ecuacion.util.excel;

import static org.assertj.core.api.Assertions.assertThat;

import java.io.IOException;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import jp.ecuacion.lib.core.util.PropertiesFileUtil;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

/**
 * Tests the message properties files of all the ecuacion-utils modules for each language.
 *
 * <p>Translations are maintained for the same languages as ecuacion-lib, and it is easy
 *     to forget to add a key or a whole file to some of them. This test reads the properties
 *     files in {@code src/main/resources} of the ecuacion-utils modules next to this module
 *     and checks that every language has the same files, keys and placeholders
 *     as the default (English) files.</p>
 */
@DisplayName("message properties files of all languages")
public class MessagePropertiesFilesTest {

  /** The languages supported by ecuacion-lib, other than English (the default). */
  private static final Set<String> LANGUAGES = Set.of("ja", "ar", "az", "cs", "da", "de", "es",
      "fa", "fr", "hu", "it", "ko", "mn", "nl", "pl", "pt", "ro", "ru", "sk", "tr", "uk", "zh",
      "zh_Hant", "id", "vi", "th", "ms", "fil", "hi", "bn", "ta", "te", "mr", "kn", "gu", "ml");

  /** Files which are not message files. */
  private static final Set<String> EXCLUDED_BASE_NAMES = Set.of("version_ecuacion-utils");

  private static final Pattern PLACEHOLDER = Pattern.compile("\\{[A-Za-z0-9_.]+}");

  /** base name -> (locale -> properties). The default file has the locale {@code ""}. */
  private static final Map<String, Map<String, Properties>> BUNDLES = new TreeMap<>();

  @SuppressWarnings("null")
  @BeforeAll
  public static void beforeAll() throws IOException {
    List<Path> resourceDirs;
    try (Stream<Path> modules = Files.list(Path.of(".."))) {
      resourceDirs = modules
          .filter(p -> p.getFileName().toString().startsWith("ecuacion-util-"))
          .map(p -> p.resolve("src/main/resources")).filter(Files::isDirectory).toList();
    }

    for (Path dir : resourceDirs) {
      try (Stream<Path> files = Files.list(dir)) {
        for (Path file : files.filter(p -> p.toString().endsWith(".properties")).toList()) {
          String name = file.getFileName().toString().replaceFirst("\\.properties$", "");
          // The locale is found from the known languages, since a base name itself can end
          // with a string like a locale (e.g. "_pdf" of "messages_util_excel_report_to_pdf").
          String locale = LANGUAGES.stream().filter(l -> name.endsWith("_" + l)).findFirst()
              .orElse("");
          String baseName =
              locale.isEmpty() ? name : name.substring(0, name.length() - locale.length() - 1);
          if (EXCLUDED_BASE_NAMES.contains(baseName)) {
            continue;
          }

          Properties props = new Properties();
          try (Reader reader = Files.newBufferedReader(file, StandardCharsets.UTF_8)) {
            props.load(reader);
          }
          Objects.requireNonNull(BUNDLES.computeIfAbsent(baseName, k -> new TreeMap<>()))
              .put(locale, props);
        }
      }
    }
  }

  @Test
  @DisplayName("message files are found")
  void filesFound() {
    assertThat(BUNDLES).containsKeys("messages_util_excel_table",
        "messages_util_excel_report_to_pdf");
  }

  @Test
  @DisplayName("every message file has the default file and the files of all the languages")
  void everyLanguageHasAllFiles() {
    Set<String> expected = new TreeSet<>(LANGUAGES);
    expected.add("");

    for (Map.Entry<String, Map<String, Properties>> bundle : BUNDLES.entrySet()) {
      assertThat(bundle.getValue().keySet()).as(bundle.getKey())
          .containsExactlyInAnyOrderElementsOf(expected);
    }
  }

  @SuppressWarnings("null")
  @Test
  @DisplayName("every language file has the same keys as the default file")
  void sameKeys() {
    List<String> errors = new ArrayList<>();
    for (Map.Entry<String, Map<String, Properties>> bundle : BUNDLES.entrySet()) {
      Properties defaultProps = Objects.requireNonNull(bundle.getValue().get(""));
      for (Map.Entry<String, Properties> entry : bundle.getValue().entrySet()) {
        Set<String> expected = defaultProps.stringPropertyNames();
        Set<String> actual = entry.getValue().stringPropertyNames();
        String file = bundle.getKey() + "_" + entry.getKey();

        Set<String> missingKeys = new TreeSet<>(expected);
        missingKeys.removeAll(actual);
        missingKeys.forEach(k -> errors.add(file + ": missing key " + k));

        Set<String> extraKeys = new TreeSet<>(actual);
        extraKeys.removeAll(expected);
        extraKeys.forEach(k -> errors.add(file + ": unknown key " + k));
      }
    }

    assertThat(errors).as("key differences").isEmpty();
  }

  @Test
  @DisplayName("every message keeps the placeholders of the default message")
  void samePlaceholders() {
    List<String> errors = new ArrayList<>();
    for (Map.Entry<String, Map<String, Properties>> bundle : BUNDLES.entrySet()) {
      Properties defaultProps = Objects.requireNonNull(bundle.getValue().get(""));
      for (Map.Entry<String, Properties> entry : bundle.getValue().entrySet()) {
        for (String key : entry.getValue().stringPropertyNames()) {
          String defaultValue = defaultProps.getProperty(key);
          if (defaultValue == null) {
            continue;
          }

          List<String> expected = placeholders(defaultValue);
          List<String> actual =
              placeholders(Objects.requireNonNull(entry.getValue().getProperty(key)));
          if (!expected.equals(actual)) {
            errors.add(bundle.getKey() + "_" + entry.getKey() + ": " + key + " has " + actual
                + " but " + expected + " is expected");
          }
        }
      }
    }

    assertThat(errors).as("placeholder differences").isEmpty();
  }

  @Test
  @DisplayName("message files do not use ASCII apostrophes")
  void noAsciiApostrophes() {
    // Messages are formatted by MessageFormat, which treats "'" as a quote,
    // so a placeholder after "'" is not replaced. Use "’" instead.
    List<String> errors = new ArrayList<>();
    for (Map.Entry<String, Map<String, Properties>> bundle : BUNDLES.entrySet()) {
      for (Map.Entry<String, Properties> entry : bundle.getValue().entrySet()) {
        for (String key : entry.getValue().stringPropertyNames()) {
          if (Objects.requireNonNull(entry.getValue().getProperty(key)).contains("'")) {
            errors.add(bundle.getKey() + (entry.getKey().isEmpty() ? "" : "_" + entry.getKey())
                + ": " + key);
          }
        }
      }
    }

    assertThat(errors).as("messages with ASCII apostrophes").isEmpty();
  }

  @SuppressWarnings("null")
  @Test
  @DisplayName("messages of this module are resolved in every language")
  void resolvedInEveryLanguage() {
    Map<String, Properties> bundle =
        Objects.requireNonNull(BUNDLES.get("messages_util_excel_table"));
    String key = "jp.ecuacion.util.excel.SheetNotExist.message";

    List<String> errors = new ArrayList<>();
    for (String language : LANGUAGES) {
      // Locale.forLanguageTag is used since "zh_Hant" has a script, not a country.
      Locale locale = Locale.forLanguageTag(language.replace('_', '-'));
      String expected = Objects.requireNonNull(bundle.get(language).getProperty(key + ".default"))
          .replace("{0}", "S");
      String actual = PropertiesFileUtil.getMessage(locale, key, "S");
      if (!expected.equals(actual)) {
        errors.add(language + ": " + actual);
      }
    }

    assertThat(errors).as("messages not resolved in the language").isEmpty();
  }

  /** Returns placeholders like {@code {0}} in sorted order. */
  private static List<String> placeholders(String value) {
    List<String> list = new ArrayList<>();
    Matcher m = PLACEHOLDER.matcher(value);
    while (m.find()) {
      list.add(m.group());
    }

    list.sort(Comparator.naturalOrder());
    return list;
  }
}
