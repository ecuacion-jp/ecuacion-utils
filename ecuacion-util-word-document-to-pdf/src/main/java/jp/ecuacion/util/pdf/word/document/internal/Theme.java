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

import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Holds the fonts of the document theme, which styles refer to by names like minorEastAsia. */
final class Theme {

  private @Nullable String majorLatin;
  private @Nullable String minorLatin;
  private @Nullable String majorEa;
  private @Nullable String minorEa;

  static Theme parse(@Nullable Element themeRoot, @Nullable String eastAsiaLang) {
    Theme t = new Theme();
    Element scheme = Xml.descendant(themeRoot, "fontScheme");
    if (scheme == null) {
      return t;
    }
    String script = script(eastAsiaLang);
    Element major = Xml.childAnyNs(scheme, "majorFont");
    Element minor = Xml.childAnyNs(scheme, "minorFont");
    t.majorLatin = typeface(Xml.childAnyNs(major, "latin"));
    t.minorLatin = typeface(Xml.childAnyNs(minor, "latin"));
    t.majorEa = eastAsia(major, script);
    t.minorEa = eastAsia(minor, script);
    return t;
  }

  private static String script(@Nullable String lang) {
    if (lang == null) {
      return "Jpan";
    }
    String l = lang.toLowerCase(java.util.Locale.ROOT);
    if (l.startsWith("zh")) {
      return l.contains("tw") || l.contains("hk") || l.contains("mo") || l.contains("hant")
          ? "Hant" : "Hans";
    }
    return l.startsWith("ko") ? "Hang" : "Jpan";
  }

  private static @Nullable String typeface(@Nullable Element e) {
    String v = Xml.attr(e, "typeface");
    return v == null || v.isBlank() ? null : v;
  }

  private static @Nullable String eastAsia(@Nullable Element fontCollection, String script) {
    String ea = typeface(Xml.childAnyNs(fontCollection, "ea"));
    if (ea != null) {
      return ea;
    }
    for (Element e : Xml.children(fontCollection)) {
      if ("font".equals(e.getLocalName()) && script.equals(Xml.attr(e, "script"))) {
        return typeface(e);
      }
    }
    return null;
  }

  /** Resolves a theme font reference (e.g. {@code minorEastAsia}) to a font name. */
  @Nullable String resolve(@Nullable String themeRef) {
    if (themeRef == null) {
      return null;
    }
    boolean major = themeRef.startsWith("major");
    if (themeRef.endsWith("EastAsia")) {
      return major ? majorEa : minorEa;
    }
    if (themeRef.endsWith("Bidi")) {
      return null;
    }
    return major ? majorLatin : minorLatin;
  }
}
