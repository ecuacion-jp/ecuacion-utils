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
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Holds paragraph properties. {@code null} means "not specified at this level". */
final class ParaProps {

  /**
   * Holds a custom tab stop.
   *
   * @param pos the position from the left edge of the text area in points
   * @param align left / center / right / decimal / clear
   * @param leader the leader kind, or {@code null}
   */
  record TabStop(float pos, String align, @Nullable String leader) {
  }

  @Nullable String styleId;
  @Nullable String jc;
  @Nullable Float indLeft;
  @Nullable Float indRight;
  /** The first line indent in points. Negative for a hanging indent. */
  @Nullable Float firstLine;
  @Nullable Integer leftChars;
  @Nullable Integer rightChars;
  /** The first line indent in 1/100 characters. Negative for a hanging indent. */
  @Nullable Integer firstLineChars;
  @Nullable Float before;
  @Nullable Float after;
  @Nullable Integer beforeLines;
  @Nullable Integer afterLines;
  /** The raw {@code w:line} value (240ths of a line for auto, twips otherwise). */
  @Nullable Float line;
  @Nullable String lineRule;
  @Nullable Boolean keepNext;
  @Nullable Boolean keepLines;
  @Nullable Boolean pageBreakBefore;
  @Nullable Boolean widowControl;
  @Nullable Boolean snapToGrid;
  @Nullable Boolean contextualSpacing;
  @Nullable Boolean autoSpaceDe;
  @Nullable Boolean autoSpaceDn;
  @Nullable Boolean wordWrap;
  @Nullable Boolean overflowPunct;
  @Nullable Integer numId;
  @Nullable Integer ilvl;
  @Nullable Integer outlineLvl;
  @Nullable Integer shade;
  Map<String, Border> borders = new HashMap<>();
  List<TabStop> tabs = new ArrayList<>();
  /** Whether {@code w:ind} is specified at this level. */
  boolean hasInd;

  ParaProps copy() {
    ParaProps c = new ParaProps();
    c.apply(this);
    return c;
  }

  /** Overrides this with the specified (non-null) values of {@code o}. */
  void apply(@Nullable ParaProps o) {
    if (o == null) {
      return;
    }
    styleId = o.styleId != null ? o.styleId : styleId;
    jc = o.jc != null ? o.jc : jc;
    applyIndent(o);
    before = o.before != null ? o.before : before;
    after = o.after != null ? o.after : after;
    beforeLines = o.beforeLines != null ? o.beforeLines : beforeLines;
    afterLines = o.afterLines != null ? o.afterLines : afterLines;
    line = o.line != null ? o.line : line;
    lineRule = o.lineRule != null ? o.lineRule : lineRule;
    keepNext = o.keepNext != null ? o.keepNext : keepNext;
    keepLines = o.keepLines != null ? o.keepLines : keepLines;
    pageBreakBefore = o.pageBreakBefore != null ? o.pageBreakBefore : pageBreakBefore;
    widowControl = o.widowControl != null ? o.widowControl : widowControl;
    snapToGrid = o.snapToGrid != null ? o.snapToGrid : snapToGrid;
    contextualSpacing = o.contextualSpacing != null ? o.contextualSpacing : contextualSpacing;
    autoSpaceDe = o.autoSpaceDe != null ? o.autoSpaceDe : autoSpaceDe;
    autoSpaceDn = o.autoSpaceDn != null ? o.autoSpaceDn : autoSpaceDn;
    wordWrap = o.wordWrap != null ? o.wordWrap : wordWrap;
    overflowPunct = o.overflowPunct != null ? o.overflowPunct : overflowPunct;
    numId = o.numId != null ? o.numId : numId;
    ilvl = o.ilvl != null ? o.ilvl : ilvl;
    outlineLvl = o.outlineLvl != null ? o.outlineLvl : outlineLvl;
    shade = o.shade != null ? o.shade : shade;
    if (!o.borders.isEmpty()) {
      borders = new HashMap<>(borders);
      borders.putAll(o.borders);
    }
    applyTabs(o);
  }

  /** Overrides the indents only. */
  void applyIndent(ParaProps o) {
    indLeft = o.indLeft != null ? o.indLeft : indLeft;
    indRight = o.indRight != null ? o.indRight : indRight;
    leftChars = o.leftChars != null ? o.leftChars : leftChars;
    rightChars = o.rightChars != null ? o.rightChars : rightChars;
    if (o.firstLine != null || o.firstLineChars != null) {
      firstLine = o.firstLine;
      firstLineChars = o.firstLineChars;
    }
    hasInd |= o.hasInd;
  }

  /** Merges the tab stops: a stop at the same position is replaced, "clear" removes it. */
  void applyTabs(ParaProps o) {
    if (o.tabs.isEmpty()) {
      return;
    }
    List<TabStop> merged = new ArrayList<>(tabs);
    for (TabStop t : o.tabs) {
      merged.removeIf(x -> Math.abs(x.pos() - t.pos()) < 0.05f);
      if (!"clear".equals(t.align())) {
        merged.add(t);
      }
    }
    merged.sort(Comparator.comparingDouble(TabStop::pos));
    tabs = merged;
  }

  static ParaProps parse(@Nullable Element ppr) {
    ParaProps p = new ParaProps();
    if (ppr == null) {
      return p;
    }
    p.styleId = Xml.val(ppr, "pStyle");
    p.jc = Xml.val(ppr, "jc");
    Element ind = Xml.child(ppr, "ind");
    if (ind != null) {
      p.hasInd = true;
      p.indLeft = first(Xml.twips(ind, "left"), Xml.twips(ind, "start"));
      p.indRight = first(Xml.twips(ind, "right"), Xml.twips(ind, "end"));
      p.leftChars = first(Xml.intAttr(ind, "leftChars"), Xml.intAttr(ind, "startChars"));
      p.rightChars = first(Xml.intAttr(ind, "rightChars"), Xml.intAttr(ind, "endChars"));
      Float hanging = Xml.twips(ind, "hanging");
      Float firstLine = Xml.twips(ind, "firstLine");
      p.firstLine = hanging != null ? Float.valueOf(-hanging) : firstLine;
      Integer hangingChars = Xml.intAttr(ind, "hangingChars");
      Integer firstLineChars = Xml.intAttr(ind, "firstLineChars");
      p.firstLineChars = hangingChars != null ? Integer.valueOf(-hangingChars) : firstLineChars;
    }
    Element spacing = Xml.child(ppr, "spacing");
    if (spacing != null) {
      p.before = Xml.twips(spacing, "before");
      p.after = Xml.twips(spacing, "after");
      p.beforeLines = Xml.intAttr(spacing, "beforeLines");
      p.afterLines = Xml.intAttr(spacing, "afterLines");
      Integer line = Xml.intAttr(spacing, "line");
      p.line = line == null ? null : Float.valueOf(line);
      p.lineRule = Xml.attr(spacing, "lineRule");
      if (p.line != null && p.lineRule == null) {
        p.lineRule = "auto";
      }
    }
    p.keepNext = Xml.onOff(ppr, "keepNext");
    p.keepLines = Xml.onOff(ppr, "keepLines");
    p.pageBreakBefore = Xml.onOff(ppr, "pageBreakBefore");
    p.widowControl = Xml.onOff(ppr, "widowControl");
    p.snapToGrid = Xml.onOff(ppr, "snapToGrid");
    p.contextualSpacing = Xml.onOff(ppr, "contextualSpacing");
    p.autoSpaceDe = Xml.onOff(ppr, "autoSpaceDE");
    p.autoSpaceDn = Xml.onOff(ppr, "autoSpaceDN");
    p.wordWrap = Xml.onOff(ppr, "wordWrap");
    p.overflowPunct = Xml.onOff(ppr, "overflowPunct");
    Element numPr = Xml.child(ppr, "numPr");
    if (numPr != null) {
      p.numId = Xml.intAttr(Xml.child(numPr, "numId"), "val");
      p.ilvl = Xml.intAttr(Xml.child(numPr, "ilvl"), "val");
    }
    p.outlineLvl = Xml.intAttr(Xml.child(ppr, "outlineLvl"), "val");
    Element shd = Xml.child(ppr, "shd");
    if (shd != null) {
      Integer fill = Xml.color(Xml.attr(shd, "fill"));
      p.shade = fill != null ? fill : RunProps.NO_COLOR;
    }
    p.borders = Border.parseAll(Xml.child(ppr, "pBdr"));
    for (Element tab : Xml.children(Xml.child(ppr, "tabs"), "tab")) {
      Float pos = Xml.twips(tab, "pos");
      String val = Xml.attr(tab, "val");
      if (pos != null) {
        String align = val == null || "start".equals(val) || "num".equals(val) ? "left"
            : "end".equals(val) ? "right" : val;
        String leader = Xml.attr(tab, "leader");
        p.tabs.add(new TabStop(pos, align, "none".equals(leader) ? null : leader));
      }
    }
    return p;
  }

  private static <T> @Nullable T first(@Nullable T a, @Nullable T b) {
    return a != null ? a : b;
  }
}
