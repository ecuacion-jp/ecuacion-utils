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

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Resolves list / heading numbering defined in numbering.xml, and counts the numbers. */
final class Numbering {

  private static final int LEVELS = 9;

  /** Holds the definition of a numbering level. */
  static final class Level {
    int start = 1;
    String format = "decimal";
    String text = "";
    String suffix = "tab";
    String justification = "left";
    boolean legal;
    /** The level whose increment restarts this level. -1: the default (any upper level). */
    int restart = -1;
    ParaProps ppr = new ParaProps();
    RunProps rpr = new RunProps();
  }

  /**
   * Holds the number label of a paragraph.
   *
   * @param text the label like "1.2" or "第3条"
   * @param rpr the run properties specified by the numbering level
   * @param suffix tab / space / nothing
   * @param justification left / center / right
   */
  record Label(String text, RunProps rpr, String suffix, String justification) {
  }

  private static final class Counter {
    final int[] value = new int[LEVELS];
    final boolean[] set = new boolean[LEVELS];
    final Integer[] pendingStart = new Integer[LEVELS];
  }

  private final Map<Integer, Level[]> abstractNums = new HashMap<>();
  private final Map<Integer, Integer> numToAbstract = new HashMap<>();
  private final Map<Integer, Level[]> overrideLevels = new HashMap<>();
  private final Map<Integer, Integer[]> startOverrides = new HashMap<>();
  private final Map<Integer, Counter> counters = new HashMap<>();
  private final Set<Integer> usedNums = new HashSet<>();

  Numbering(@Nullable Element root, Theme theme) {
    Map<String, Integer> styleLinks = new HashMap<>();
    Map<Integer, String> numStyleLinks = new HashMap<>();
    for (Element an : Xml.children(root, "abstractNum")) {
      Integer id = Xml.intAttr(an, "abstractNumId");
      if (id == null) {
        continue;
      }
      Level[] levels = new Level[LEVELS];
      for (Element lvl : Xml.children(an, "lvl")) {
        Integer ilvl = Xml.intAttr(lvl, "ilvl");
        if (ilvl != null && ilvl >= 0 && ilvl < LEVELS) {
          levels[ilvl] = parseLevel(lvl, theme);
        }
      }
      abstractNums.put(id, levels);
      String styleLink = Xml.val(an, "styleLink");
      if (styleLink != null) {
        styleLinks.put(styleLink, id);
      }
      String numStyleLink = Xml.val(an, "numStyleLink");
      if (numStyleLink != null) {
        numStyleLinks.put(id, numStyleLink);
      }
    }
    // An abstractNum with numStyleLink has no levels of its own: it uses those of the
    // abstractNum which declares the same name as its styleLink.
    for (Map.Entry<Integer, String> e : numStyleLinks.entrySet()) {
      Integer target = styleLinks.get(e.getValue());
      if (target != null) {
        abstractNums.put(e.getKey(), abstractNums.get(target));
      }
    }
    for (Element num : Xml.children(root, "num")) {
      Integer numId = Xml.intAttr(num, "numId");
      Integer abs = Xml.intAttr(Xml.child(num, "abstractNumId"), "val");
      if (numId == null || abs == null) {
        continue;
      }
      numToAbstract.put(numId, abs);
      for (Element ov : Xml.children(num, "lvlOverride")) {
        Integer ilvl = Xml.intAttr(ov, "ilvl");
        if (ilvl == null || ilvl < 0 || ilvl >= LEVELS) {
          continue;
        }
        Integer start = Xml.intAttr(Xml.child(ov, "startOverride"), "val");
        if (start != null) {
          Integer[] starts = Objects.requireNonNull(
              startOverrides.computeIfAbsent(numId, k -> new Integer[LEVELS]));
          starts[ilvl] = start;
        }
        Element lvl = Xml.child(ov, "lvl");
        if (lvl != null) {
          Level[] levels =
              Objects.requireNonNull(overrideLevels.computeIfAbsent(numId, k -> new Level[LEVELS]));
          levels[ilvl] = parseLevel(lvl, theme);
        }
      }
    }
  }

  private static Level parseLevel(Element lvl, Theme theme) {
    Level l = new Level();
    Integer start = Xml.intAttr(Xml.child(lvl, "start"), "val");
    l.start = start == null ? 1 : start;
    String fmt = Xml.val(lvl, "numFmt");
    l.format = fmt == null ? "decimal" : fmt;
    String text = Xml.val(lvl, "lvlText");
    l.text = text == null ? "" : text;
    String suff = Xml.val(lvl, "suff");
    l.suffix = suff == null ? "tab" : suff;
    String jc = Xml.val(lvl, "lvlJc");
    l.justification = jc == null || "start".equals(jc) ? "left" : "end".equals(jc) ? "right" : jc;
    l.legal = Boolean.TRUE.equals(Xml.onOff(lvl, "isLgl"));
    Integer restart = Xml.intAttr(Xml.child(lvl, "lvlRestart"), "val");
    l.restart = restart == null ? -1 : restart;
    l.ppr = ParaProps.parse(Xml.child(lvl, "pPr"));
    l.rpr = RunProps.parse(Xml.child(lvl, "rPr"), theme);
    return l;
  }

  /** Returns the level definition, or {@code null} if the numbering does not exist. */
  @Nullable Level level(int numId, int ilvl) {
    if (ilvl < 0 || ilvl >= LEVELS) {
      return null;
    }
    Level[] ov = overrideLevels.get(numId);
    if (ov != null && ov[ilvl] != null) {
      return ov[ilvl];
    }
    Integer abs = numToAbstract.get(numId);
    Level[] levels = abs == null ? null : abstractNums.get(abs);
    return levels == null ? null : levels[ilvl];
  }

  /** Advances the counter of the level and returns the label of the paragraph. */
  @Nullable Label next(int numId, int ilvl) {
    Level level = level(numId, ilvl);
    Integer abs = numToAbstract.get(numId);
    if (level == null || abs == null) {
      return null;
    }
    Counter c = Objects.requireNonNull(counters.computeIfAbsent(abs, k -> new Counter()));
    if (usedNums.add(numId)) {
      Integer[] starts = startOverrides.get(numId);
      if (starts != null) {
        for (int i = 0; i < LEVELS; i++) {
          if (starts[i] != null) {
            c.set[i] = false;
            c.pendingStart[i] = starts[i];
          }
        }
      }
    }
    for (int i = 0; i < ilvl; i++) {
      if (!c.set[i]) {
        c.value[i] = startOf(c, numId, i);
        c.set[i] = true;
      }
    }
    if (c.set[ilvl]) {
      c.value[ilvl]++;
    } else {
      c.value[ilvl] = startOf(c, numId, ilvl);
      c.set[ilvl] = true;
    }
    for (int i = ilvl + 1; i < LEVELS; i++) {
      Level lower = level(numId, i);
      // lvlRestart 0: never restarts. n > 0: restarts only when level n (1-based) or an
      // upper level is incremented.
      boolean restarts = lower == null || lower.restart < 0
          || (lower.restart > 0 && ilvl + 1 <= lower.restart);
      if (restarts) {
        c.set[i] = false;
      }
    }

    String text;
    if ("bullet".equals(level.format)) {
      text = bullet(level.text);
    } else {
      StringBuilder sb = new StringBuilder();
      String t = level.text;
      for (int i = 0; i < t.length(); i++) {
        char ch = t.charAt(i);
        if (ch == '%' && i + 1 < t.length() && t.charAt(i + 1) >= '1'
            && t.charAt(i + 1) <= '9') {
          int ref = t.charAt(i + 1) - '1';
          Level refLevel = level(numId, ref);
          String fmt = refLevel == null ? "decimal" : refLevel.format;
          if (level.legal && ref != ilvl) {
            fmt = "decimal";
          }
          sb.append(NumberFormats.format(c.set[ref] ? c.value[ref] : 0, fmt));
          i++;
        } else {
          sb.append(ch);
        }
      }
      text = sb.toString();
    }
    return new Label(text, level.rpr, level.suffix, level.justification);
  }

  private int startOf(Counter c, int numId, int ilvl) {
    Integer pending = c.pendingStart[ilvl];
    if (pending != null) {
      c.pendingStart[ilvl] = null;
      return pending;
    }
    Level l = level(numId, ilvl);
    return l == null ? 1 : l.start;
  }

  /** Maps the bullets of symbol fonts (private use area) to Unicode characters. */
  private static String bullet(String text) {
    StringBuilder sb = new StringBuilder();
    for (int i = 0; i < text.length(); i++) {
      char ch = text.charAt(i);
      sb.append(switch (ch) {
        case '', '·' -> '•';
        case '' -> '▪';
        case '' -> '●';
        case '' -> '■';
        case '' -> '◆';
        case '' -> '➢';
        case '' -> '✓';
        case '' -> '◆';
        default -> ch >= '' && ch <= '' ? '•' : ch;
      });
    }
    return sb.toString();
  }
}
