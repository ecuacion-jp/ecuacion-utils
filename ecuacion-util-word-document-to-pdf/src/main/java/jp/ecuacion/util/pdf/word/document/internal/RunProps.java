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

import java.util.Map;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Holds run (character) properties. {@code null} means "not specified at this level". */
final class RunProps {

  private static final Map<String, Integer> HIGHLIGHT_COLORS = Map.ofEntries(
      Map.entry("yellow", 0xFFFF00), Map.entry("green", 0x00FF00), Map.entry("cyan", 0x00FFFF),
      Map.entry("magenta", 0xFF00FF), Map.entry("blue", 0x0000FF), Map.entry("red", 0xFF0000),
      Map.entry("darkBlue", 0x000080), Map.entry("darkCyan", 0x008080),
      Map.entry("darkGreen", 0x008000), Map.entry("darkMagenta", 0x800080),
      Map.entry("darkRed", 0x800000), Map.entry("darkYellow", 0x808000),
      Map.entry("darkGray", 0x808080), Map.entry("lightGray", 0xC0C0C0),
      Map.entry("black", 0x000000), Map.entry("white", 0xFFFFFF));

  /** The value which means "no highlight / shading", to override an inherited one. */
  static final int NO_COLOR = -1;

  @Nullable String ascii;
  @Nullable String hansi;
  @Nullable String eastAsia;
  @Nullable String hint;
  @Nullable Boolean bold;
  @Nullable Boolean italic;
  @Nullable Boolean strike;
  @Nullable Boolean caps;
  @Nullable Boolean vanish;
  /** The underline style ({@code none} for no underline). */
  @Nullable String underline;
  /** RGB, or {@link #NO_COLOR} for {@code auto}. */
  @Nullable Integer color;
  /** The font size in points. */
  @Nullable Float size;
  @Nullable Integer highlight;
  @Nullable Integer shade;
  @Nullable String vertAlign;
  /** The character spacing in points. */
  @Nullable Float spacing;
  @Nullable String styleId;

  /** The cache key of the font names and the weight, built lazily. */
  private @Nullable String fontKey;

  RunProps copy() {
    RunProps c = new RunProps();
    c.apply(this);
    return c;
  }

  /** Overrides this with the specified (non-null) values of {@code o}. */
  void apply(@Nullable RunProps o) {
    if (o == null) {
      return;
    }
    ascii = o.ascii != null ? o.ascii : ascii;
    hansi = o.hansi != null ? o.hansi : hansi;
    eastAsia = o.eastAsia != null ? o.eastAsia : eastAsia;
    hint = o.hint != null ? o.hint : hint;
    bold = o.bold != null ? o.bold : bold;
    italic = o.italic != null ? o.italic : italic;
    strike = o.strike != null ? o.strike : strike;
    caps = o.caps != null ? o.caps : caps;
    vanish = o.vanish != null ? o.vanish : vanish;
    underline = o.underline != null ? o.underline : underline;
    color = o.color != null ? o.color : color;
    size = o.size != null ? o.size : size;
    highlight = o.highlight != null ? o.highlight : highlight;
    shade = o.shade != null ? o.shade : shade;
    vertAlign = o.vertAlign != null ? o.vertAlign : vertAlign;
    spacing = o.spacing != null ? o.spacing : spacing;
    styleId = o.styleId != null ? o.styleId : styleId;
    fontKey = null;
  }

  boolean isBold() {
    return Boolean.TRUE.equals(bold);
  }

  float sizeOrDefault() {
    Float s = size;
    return s != null ? s : 10f;
  }

  String fontKey() {
    String k = fontKey;
    if (k == null) {
      k = ascii + "|" + hansi + "|" + eastAsia + "|" + hint + "|" + isBold();
      fontKey = k;
    }
    return k;
  }

  static RunProps parse(@Nullable Element rpr, Theme theme) {
    RunProps p = new RunProps();
    if (rpr == null) {
      return p;
    }
    Element fonts = Xml.child(rpr, "rFonts");
    if (fonts != null) {
      p.ascii = font(fonts, "ascii", "asciiTheme", theme);
      p.hansi = font(fonts, "hAnsi", "hAnsiTheme", theme);
      p.eastAsia = font(fonts, "eastAsia", "eastAsiaTheme", theme);
      p.hint = Xml.attr(fonts, "hint");
    }
    p.styleId = Xml.val(rpr, "rStyle");
    p.bold = Xml.onOff(rpr, "b");
    p.italic = Xml.onOff(rpr, "i");
    Boolean strike = Xml.onOff(rpr, "strike");
    Boolean dstrike = Xml.onOff(rpr, "dstrike");
    p.strike = strike != null ? strike : dstrike;
    p.caps = Xml.onOff(rpr, "caps");
    p.vanish = Xml.onOff(rpr, "vanish");
    p.underline = Xml.val(rpr, "u");
    String color = Xml.val(rpr, "color");
    if (color != null) {
      Integer rgb = Xml.color(color);
      p.color = rgb != null ? rgb : NO_COLOR;
    }
    Element sz = Xml.child(rpr, "sz");
    Integer halfPoints = Xml.intAttr(sz, "val");
    if (halfPoints != null) {
      p.size = halfPoints / 2f;
    }
    String highlight = Xml.val(rpr, "highlight");
    if (highlight != null) {
      p.highlight = HIGHLIGHT_COLORS.getOrDefault(highlight, NO_COLOR);
    }
    Element shd = Xml.child(rpr, "shd");
    if (shd != null) {
      Integer fill = Xml.color(Xml.attr(shd, "fill"));
      p.shade = fill != null ? fill : NO_COLOR;
    }
    p.vertAlign = Xml.val(rpr, "vertAlign");
    p.spacing = Xml.twips(Xml.child(rpr, "spacing"), "val");
    return p;
  }

  private static @Nullable String font(Element fonts, String attr, String themeAttr,
      Theme theme) {
    String themed = theme.resolve(Xml.attr(fonts, themeAttr));
    if (themed != null) {
      return themed;
    }
    String name = Xml.attr(fonts, attr);
    return name == null || name.isBlank() ? null : name;
  }
}
