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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Resolves the style hierarchy of styles.xml (document defaults and "based on" chains). */
final class StyleSheet {

  private record Style(String id, @Nullable String basedOn, Element element) {
  }

  /** Holds the formatting a table style applies to a part of a table (or to the whole). */
  static final class TablePart {
    ParaProps ppr = new ParaProps();
    RunProps rpr = new RunProps();
    Map<String, Border> cellBorders = new HashMap<>();
    @Nullable Integer shade;
  }

  /** Holds a resolved table style. */
  static final class TableStyle {
    TablePart whole = new TablePart();
    Map<String, Border> borders = new HashMap<>();
    /** left, top, right, bottom in points. */
    @Nullable Float[] cellMargins = new Float[4];
    Map<String, TablePart> conditions = new HashMap<>();
  }

  private final Map<String, Style> styles = new HashMap<>();
  private final Theme theme;
  private final RunProps defaultRun;
  private final ParaProps defaultPara;
  private @Nullable String defaultParaStyle;
  private @Nullable String defaultTableStyle;
  private final Map<String, TableStyle> tableStyles = new HashMap<>();

  StyleSheet(@Nullable Element stylesRoot, Theme theme) {
    this.theme = theme;
    Element docDefaults = Xml.child(stylesRoot, "docDefaults");
    defaultRun =
        RunProps.parse(Xml.child(Xml.child(docDefaults, "rPrDefault"), "rPr"), theme);
    defaultPara = ParaProps.parse(Xml.child(Xml.child(docDefaults, "pPrDefault"), "pPr"));
    for (Element s : Xml.children(stylesRoot, "style")) {
      String id = Xml.attr(s, "styleId");
      if (id == null) {
        continue;
      }
      styles.put(id, new Style(id, Xml.val(s, "basedOn"), s));
      String isDefault = Xml.attr(s, "default");
      if ("1".equals(isDefault) || "true".equals(isDefault)) {
        String type = Xml.attr(s, "type");
        if ("paragraph".equals(type)) {
          defaultParaStyle = id;
        } else if ("table".equals(type)) {
          defaultTableStyle = id;
        }
      }
    }
  }

  Theme theme() {
    return theme;
  }

  @Nullable String defaultParaStyle() {
    return defaultParaStyle;
  }

  /** Returns the style chain, the base style first. */
  private List<Style> chain(@Nullable String styleId) {
    List<Style> list = new ArrayList<>();
    Set<String> seen = new HashSet<>();
    String id = styleId;
    while (id != null && seen.add(id)) {
      Style s = styles.get(id);
      if (s == null) {
        break;
      }
      list.add(0, s);
      id = s.basedOn();
    }
    return list;
  }

  /**
   * Holds the paragraph properties from the styles, and whether the indent of the numbering
   * level referred to by the style takes effect (it does unless a style at the same or a more
   * derived level specifies an indent itself).
   */
  record ParaBase(ParaProps props, boolean numberingIndentApplies) {
  }

  /** Returns document defaults + table style + paragraph style chain. */
  ParaBase paraBase(@Nullable String styleId, @Nullable ParaProps tableProps) {
    ParaProps p = defaultPara.copy();
    p.hasInd = false;
    p.apply(tableProps);
    int indIndex = -1;
    int numIndex = -1;
    List<Style> chain = chain(styleId != null ? styleId : defaultParaStyle);
    for (int i = 0; i < chain.size(); i++) {
      ParaProps sp = ParaProps.parse(Xml.child(chain.get(i).element(), "pPr"));
      if (sp.hasInd) {
        indIndex = i;
      }
      if (sp.numId != null) {
        numIndex = i;
      }
      p.apply(sp);
    }
    return new ParaBase(p, numIndex > indIndex);
  }

  /** Returns document defaults + table style + the run properties of the paragraph style. */
  RunProps runBase(@Nullable String paraStyleId, @Nullable RunProps tableProps) {
    RunProps r = defaultRun.copy();
    r.apply(tableProps);
    for (Style s : chain(paraStyleId != null ? paraStyleId : defaultParaStyle)) {
      r.apply(RunProps.parse(Xml.child(s.element(), "rPr"), theme));
    }
    return r;
  }

  /** Returns the run properties of a character style chain (without document defaults). */
  RunProps charStyle(@Nullable String styleId) {
    RunProps r = new RunProps();
    for (Style s : chain(styleId)) {
      r.apply(RunProps.parse(Xml.child(s.element(), "rPr"), theme));
    }
    return r;
  }

  /** Returns the resolved table style ({@code null} id means the default table style). */
  TableStyle tableStyle(@Nullable String styleId) {
    String id = styleId != null ? styleId : defaultTableStyle;
    if (id == null) {
      return new TableStyle();
    }
    return Objects.requireNonNull(tableStyles.computeIfAbsent(id, this::resolveTableStyle));
  }

  private TableStyle resolveTableStyle(String id) {
    TableStyle ts = new TableStyle();
    for (Style s : chain(id)) {
      Element e = s.element();
      applyPart(ts.whole, e);
      Element tblPr = Xml.child(e, "tblPr");
      ts.borders.putAll(Border.parseAll(Xml.child(tblPr, "tblBorders")));
      applyCellMargins(ts.cellMargins, Xml.child(tblPr, "tblCellMar"));
      for (Element cond : Xml.children(e, "tblStylePr")) {
        String type = Xml.attr(cond, "type");
        if (type != null) {
          TablePart part =
              Objects.requireNonNull(ts.conditions.computeIfAbsent(type, k -> new TablePart()));
          applyPart(part, cond);
        }
      }
    }
    return ts;
  }

  private void applyPart(TablePart part, Element e) {
    part.ppr.apply(ParaProps.parse(Xml.child(e, "pPr")));
    part.rpr.apply(RunProps.parse(Xml.child(e, "rPr"), theme));
    Element tcPr = Xml.child(e, "tcPr");
    part.cellBorders.putAll(Border.parseAll(Xml.child(tcPr, "tcBorders")));
    Element shd = Xml.child(tcPr, "shd");
    if (shd != null) {
      Integer fill = Xml.color(Xml.attr(shd, "fill"));
      part.shade = fill != null ? fill : RunProps.NO_COLOR;
    }
  }

  /** Reads left / top / right / bottom cell margins into the array, where specified. */
  static void applyCellMargins(@Nullable Float[] margins, @Nullable Element mar) {
    if (mar == null) {
      return;
    }
    String[][] names = {{"left", "start"}, {"top"}, {"right", "end"}, {"bottom"}};
    for (int i = 0; i < 4; i++) {
      for (String name : names[i]) {
        Float v = Xml.twips(Xml.child(mar, name), "w");
        if (v != null) {
          margins[i] = v;
        }
      }
    }
  }
}
