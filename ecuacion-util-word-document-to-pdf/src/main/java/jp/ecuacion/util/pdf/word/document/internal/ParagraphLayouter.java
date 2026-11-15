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
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Breaks a paragraph into lines and positions the text of each line. */
final class ParagraphLayouter {

  private static final int CHAR = 0;
  private static final int TAB = 1;
  private static final int IMAGE = 2;
  private static final int BREAK = 3;
  private static final int PAGE_BREAK = 4;
  private static final float EPS = 0.01f;
  /** The margin of the limit of compressing punctuation (half a character), about 2 twips. */
  private static final float FIT_TOLERANCE = 0.1f;
  /** The additional advance of emboldened text, as a ratio of the font size. */
  private static final float SYNTHETIC_BOLD_ADVANCE = 0.02f;
  /** The ratio by which a space between words can shrink in a justified line. */
  private static final float SPACE_SHRINK = 0.25f;
  private static final Layout.Style PLAIN = new Layout.Style(0, false, false, null, false, null);

  /** An unbreakable unit of a paragraph: a character, a tab, a picture or a break. */
  private static final class Item {
    final int kind;
    /** The character for line breaking rules. */
    int cp;
    /** The character actually drawn. */
    int outCp;
    FontResolver.@Nullable Face face;
    /** The font size of the run. */
    float size;
    /** The size the glyph is drawn at (smaller for superscript / subscript). */
    float drawSize;
    /** The advance of the glyph itself. */
    float natural;
    /** The advance including spacing adjustments. For a tab, its width on the line. */
    float adv;
    float rise;
    Layout.Style style = PLAIN;
    RunProps rpr = new RunProps();
    Model.@Nullable Image image;
    @Nullable String absoluteTab;
    @Nullable String leader;
    /** The PAGE field the item is a part of. */
    Model.@Nullable Field pageField;
    /** Whether the advance has already been reduced by punctuation compression. */
    boolean compressed;
    /** The automatic space between East Asian and Latin text added after this item. */
    float autoSpace;

    Item(int kind) {
      this.kind = kind;
    }

    boolean space() {
      return kind == CHAR && cp == ' ';
    }
  }

  private record TabTarget(float pos, String align, @Nullable String leader) {
  }

  private final FontResolver fonts;
  private final Model.Settings settings;

  ParagraphLayouter(FontResolver fonts, Model.Settings settings) {
    this.fonts = fonts;
    this.settings = settings;
  }

  /** The items of a paragraph, and what is attached to them. */
  private static final class Prepared {
    final List<Item> items = new ArrayList<>();
    /** item index -> the bookmarks which start there. */
    final Map<Integer, List<String>> bookmarks = new HashMap<>();
    /** item index -> the floating pictures anchored there. */
    final Map<Integer, List<Model.Image>> floating = new HashMap<>();
    final StringBuilder plain = new StringBuilder();
    float labelShift;
  }

  private Prepared prepare(Model.Para para, Layout.Env env) {
    final Prepared prepared = new Prepared();
    final ParaProps pp = para.props;
    final List<Item> items = prepared.items;
    final Map<Integer, List<String>> bookmarks = prepared.bookmarks;
    final Map<Integer, List<Model.Image>> floating = prepared.floating;
    final StringBuilder plain = prepared.plain;

    float labelShift = 0f;
    Numbering.Label label = para.label;
    RunProps labelProps = para.labelProps;
    if (label != null && labelProps != null
        && !(label.text().isEmpty() && "nothing".equals(label.suffix()))) {
      addText(items, label.text(), labelProps, null, env);
      float labelWidth = 0f;
      for (Item it : items) {
        labelWidth += it.adv;
      }
      if ("right".equals(label.justification())) {
        labelShift = -labelWidth;
      } else if ("center".equals(label.justification())) {
        labelShift = -labelWidth / 2;
      }
      if ("tab".equals(label.suffix())) {
        Item tab = new Item(TAB);
        tab.rpr = labelProps;
        tab.size = labelProps.sizeOrDefault();
        items.add(tab);
      } else if ("space".equals(label.suffix())) {
        addText(items, " ", labelProps, null, env);
      }
      plain.append(label.text()).append(' ');
    }

    for (Model.Inline inline : para.inlines) {
      if (inline instanceof Model.Text t) {
        addText(items, t.text(), t.rpr(), t.link(), env);
        plain.append(t.text());
      } else if (inline instanceof Model.Field f) {
        int from = items.size();
        addText(items, fieldText(f, env), f.rpr(), f.link(), env);
        if ("PAGE".equals(f.kind())) {
          for (int k = from; k < items.size(); k++) {
            items.get(k).pageField = f;
          }
        }
      } else if (inline instanceof Model.Tab t) {
        Item tab = new Item(TAB);
        tab.rpr = t.rpr();
        tab.size = t.rpr().sizeOrDefault();
        tab.style = style(t.rpr(), null);
        tab.absoluteTab = t.absolute();
        items.add(tab);
        plain.append(' ');
      } else if (inline instanceof Model.Break b) {
        items.add(new Item(b.page() ? PAGE_BREAK : BREAK));
      } else if (inline instanceof Model.Image image) {
        if (image.anchor() != null) {
          Objects.requireNonNull(floating.computeIfAbsent(items.size(), k -> new ArrayList<>()))
              .add(image);
        } else {
          Item it = new Item(IMAGE);
          it.image = image;
          it.adv = image.width();
          items.add(it);
        }
      } else if (inline instanceof Model.Bookmark b) {
        Objects.requireNonNull(bookmarks.computeIfAbsent(items.size(), k -> new ArrayList<>()))
            .add(b.name());
      }
    }
    adjustSpacing(items, pp);
    prepared.labelShift = labelShift;
    return prepared;
  }

  /**
   * Returns the width the paragraph needs at least: the width of the longest part which
   * cannot be broken into lines, plus the indents.
   */
  float minContentWidth(Model.Para para, Layout.Env env) {
    final ParaProps pp = para.props;
    final List<Item> items = prepare(para, env).items;
    final boolean wordWrap = !Boolean.FALSE.equals(pp.wordWrap);
    float max = 0f;
    float current = 0f;
    for (int i = 0; i < items.size(); i++) {
      Item it = items.get(i);
      if (it.kind != CHAR && it.kind != IMAGE) {
        current = 0f;
        continue;
      }
      if (it.space() || (i > 0 && breakable(items.get(i - 1), it, wordWrap))) {
        current = 0f;
      }
      if (!it.space()) {
        current += it.adv;
        max = Math.max(max, current);
      }
    }
    return max + Math.max(0f, nz(pp.indLeft)) + Math.max(0f, nz(pp.indRight));
  }

  List<Layout.Line> layout(Model.Para para, float areaWidth, Layout.Env env) {
    final ParaProps pp = para.props;
    final Prepared prepared = prepare(para, env);
    final List<Item> items = prepared.items;
    final Map<Integer, List<String>> bookmarks = prepared.bookmarks;
    final Map<Integer, List<Model.Image>> floating = prepared.floating;
    final StringBuilder plain = prepared.plain;
    final float labelShift = prepared.labelShift;

    final float unit = env.charUnit;
    final float right = length(pp.rightChars, unit, pp.indRight);
    final float first = length(pp.firstLineChars, unit, pp.firstLine);
    final Integer leftChars = pp.leftChars;
    final Integer firstChars = pp.firstLineChars;
    // When both are given in characters, a hanging indent is added to the left indent:
    // the left indent in characters is the position of the first line.
    final float left = length(leftChars, unit, pp.indLeft)
        + (leftChars != null && leftChars != 0 && firstChars != null && firstChars < 0
            ? -first : 0f);
    final boolean wordWrap = !Boolean.FALSE.equals(pp.wordWrap);
    final boolean hanging = !Boolean.FALSE.equals(pp.overflowPunct);
    final String jc = pp.jc == null ? "left" : pp.jc;
    // Punctuation is compressed to put more on a line only in justified paragraphs.
    final boolean squeeze = "both".equals(jc) || "distribute".equals(jc);

    List<Layout.Line> lines = new ArrayList<>();
    final int n = items.size();
    int pos = 0;
    boolean firstLine = true;
    boolean again;
    do {
      final float start = left + (firstLine ? first + labelShift : 0f);
      final float avail = Math.max(areaWidth - right - start, 1f);
      float[] xs = new float[n - pos + 1];
      float x = 0f;
      int i = pos;
      int breakBefore = -1;
      int end = -1;
      boolean hard = false;
      boolean pageBreak = false;
      boolean hung = false;
      int wordAcceptedUntil = -1;
      while (i < n) {
        Item it = items.get(i);
        if (it.kind >= BREAK) {
          hard = true;
          pageBreak = it.kind == PAGE_BREAK;
          end = i;
          break;
        }
        if (i > pos && breakable(items.get(i - 1), it, wordWrap)) {
          breakBefore = i;
        }
        if (it.kind == TAB) {
          it.adv = tabWidth(items, i, pp, start + x, left, first, right, areaWidth);
        }
        if (x + it.adv > avail + EPS && i > pos && !it.space()) {
          if (i < wordAcceptedUntil) {
            xs[i - pos] = x;
            x += it.adv;
            i++;
            continue;
          }
          if (squeeze && it.kind == CHAR && breakBefore > pos
              && items.get(breakBefore - 1).space()) {
            int wordEnd = wordEnd(items, i, wordWrap);
            if (fitsByShrinkingSpaces(items, pos, breakBefore, i, wordEnd, xs, x, avail)) {
              wordAcceptedUntil = wordEnd;
              xs[i - pos] = x;
              x += it.adv;
              i++;
              continue;
            }
          }
          if (squeeze && it.kind == CHAR
              && x + it.adv - avail <= squeezable(items, pos, i, it, items.get(i - 1))) {
            // Word makes room for the character by compressing the punctuation of the line.
            xs[i - pos] = x;
            x += it.adv;
            i++;
            continue;
          }
          boolean breakableAfter = i + 1 >= n || items.get(i + 1).kind >= BREAK
              || breakable(it, items.get(i + 1), wordWrap);
          if (hanging && it.kind == CHAR && TextRules.hangable(it.cp) && breakableAfter
              && x <= avail + EPS) {
            // The punctuation hangs out of the right edge.
            xs[i - pos] = x;
            x += it.adv;
            i++;
            end = i;
            hung = true;
          } else {
            end = breakBefore > pos ? breakBefore : i;
          }
          break;
        }
        xs[i - pos] = x;
        x += it.adv;
        i++;
      }
      if (end < 0) {
        end = n;
      }
      float endX = end < i ? xs[end - pos] : x;
      final int consumed = hard ? end + 1 : end;
      final boolean lastLine = end >= n || hard;

      int last = end - 1;
      while (last >= pos && items.get(last).space()) {
        last--;
      }
      float visible = last < pos ? 0f : xs[last - pos] + items.get(last).adv;

      if (!hung && visible > avail + EPS && last > pos) {
        // The line is too long by the character squeezed in: compress the punctuation.
        float capacity = 0f;
        for (int j = pos; j < last; j++) {
          capacity += shrinkable(items.get(j));
        }
        if (capacity > 0f) {
          float ratio = Math.min(1f, (visible - avail) / capacity);
          float shift = 0f;
          for (int j = pos; j < end; j++) {
            Item it = items.get(j);
            float amount = j < last ? shrinkable(it) * ratio : 0f;
            // An opening bracket has its blank on the left, the others on the right.
            boolean before = it.autoSpace == 0f && TextRules.fullWidthOpening(it.cp);
            if (before) {
              shift += amount;
            }
            xs[j - pos] -= shift;
            if (!before) {
              shift += amount;
            }
          }
          endX -= shift;
          visible -= shift;
        }
      }

      boolean distribute = "distribute".equals(jc);
      if (("both".equals(jc) || distribute) && (!lastLine || distribute) && last > pos) {
        float extra = avail - (hung ? xs[last - pos] : visible);
        int justifyFrom = pos;
        for (int j = pos; j <= last; j++) {
          if (items.get(j).kind == TAB) {
            justifyFrom = j + 1;
          }
        }
        int lastGap = hung ? last - 1 : last;
        int gaps = 0;
        for (int j = justifyFrom; j < lastGap; j++) {
          if (gapAfter(items.get(j), items.get(j + 1))) {
            gaps++;
          }
        }
        if (extra > EPS && gaps > 0) {
          float add = extra / gaps;
          float shift = 0f;
          for (int j = justifyFrom; j < end; j++) {
            xs[j - pos] += shift;
            if (j < lastGap && gapAfter(items.get(j), items.get(j + 1))) {
              shift += add;
            }
          }
          endX += shift;
          visible += shift;
        }
      }

      Layout.Line line = new Layout.Line();
      float offset = 0f;
      if ("center".equals(jc)) {
        offset = (avail - visible) / 2;
      } else if ("right".equals(jc) || "end".equals(jc)) {
        offset = avail - visible;
      }
      line.x0 = start + offset;
      buildSegs(line, items, pos, end, xs, endX);
      measure(line, items, pos, end, para, pp, env);
      line.pageBreakAfter = pageBreak;
      line.boxLeft = left + Math.min(0f, first);
      line.boxRight = areaWidth - right;
      Integer shade = pp.shade;
      line.background = shade != null && shade >= 0 ? shade : null;
      line.borderLeft = pp.borders.get("left");
      line.borderRight = pp.borders.get("right");
      for (int j = pos; j < consumed; j++) {
        attach(line, bookmarks.get(j), floating.get(j));
        Model.Field pageField = items.get(j).pageField;
        if (pageField != null && !line.pageFields.contains(pageField)) {
          line.pageFields.add(pageField);
        }
      }
      lines.add(line);

      again = hard && consumed == n && (!pageBreak || settings.compatibilityMode < 15);
      pos = consumed;
      firstLine = false;
    } while (pos < n || again);

    Layout.Line head = lines.get(0);
    Layout.Line tail = lines.get(lines.size() - 1);
    attach(tail, bookmarks.get(n), floating.get(n));
    float lineUnit = env.linePitch > 0 ? env.linePitch : 12f;
    head.spaceBefore = length(pp.beforeLines, lineUnit, pp.before);
    tail.spaceAfter = length(pp.afterLines, lineUnit, pp.after);
    Border top = pp.borders.get("top");
    if (top != null && top.visible()) {
      head.borderTop = top;
      head.padTop = top.space() + top.width();
    }
    Border bottom = pp.borders.get("bottom");
    if (bottom != null && bottom.visible()) {
      tail.borderBottom = bottom;
      tail.padBottom = bottom.space() + bottom.width();
    }
    Integer outlineLevel = pp.outlineLvl;
    String title = plain.toString().trim();
    if (outlineLevel != null && outlineLevel >= 0 && outlineLevel < 9 && !title.isEmpty()) {
      head.outlineTitle = title;
      head.outlineLevel = outlineLevel;
    }
    return lines;
  }

  private static void attach(Layout.Line line, @Nullable List<String> bookmarks,
      @Nullable List<Model.Image> floating) {
    if (bookmarks != null) {
      for (String b : bookmarks) {
        if (!line.bookmarks.contains(b)) {
          line.bookmarks.add(b);
        }
      }
    }
    if (floating != null) {
      for (Model.Image image : floating) {
        if (!line.floating.contains(image)) {
          line.floating.add(image);
        }
      }
    }
  }

  /** Returns the length given in 1/100 units when specified, the length in points if not. */
  private static float length(@Nullable Integer hundredths, float unit, @Nullable Float points) {
    return hundredths != null && hundredths != 0 ? hundredths / 100f * unit : nz(points);
  }

  private static float nz(@Nullable Float value) {
    return value == null ? 0f : value;
  }

  private static String fieldText(Model.Field f, Layout.Env env) {
    return switch (f.kind()) {
      case "PAGE" -> {
        // In the body, the page a line is on is known only after the layout: the page of
        // the previous pass is used.
        Layout.PageCanvas known = env.fieldPages.get(f);
        int number = known != null ? known.number : env.pageNumber;
        String text = known != null ? known.numberText : env.pageText;
        yield f.format().isEmpty() ? text : NumberFormats.format(number, f.format());
      }
      case "PAGEREF" -> {
        String page = env.bookmarkPages.get(f.arg());
        yield page != null ? page : f.fallback();
      }
      case "SECTIONPAGES" -> env.sectionPages.isEmpty() ? f.fallback() : env.sectionPages;
      default -> env.numPages.isEmpty() ? f.fallback() : env.numPages;
    };
  }

  private static Layout.Style style(RunProps rp, @Nullable String link) {
    Integer color = rp.color;
    final Integer highlightColor = rp.highlight;
    final Integer shadeColor = rp.shade;
    Integer highlight = highlightColor != null && highlightColor >= 0 ? highlightColor
        : shadeColor != null && shadeColor >= 0 ? shadeColor : null;
    return new Layout.Style(color != null && color >= 0 ? color : 0,
        rp.underline != null && !"none".equals(rp.underline), Boolean.TRUE.equals(rp.strike),
        highlight, Boolean.TRUE.equals(rp.italic), link);
  }

  private void addText(List<Item> items, String raw, RunProps rp, @Nullable String link,
      Layout.Env env) {
    final String text = Boolean.TRUE.equals(rp.caps) ? raw.toUpperCase(Locale.ROOT) : raw;
    final float size = rp.sizeOrDefault();
    float drawSize = size;
    float rise = 0f;
    if ("superscript".equals(rp.vertAlign)) {
      drawSize = size * 0.65f;
      // Measured on a PDF file exported by Word.
      rise = size * 0.215f;
    } else if ("subscript".equals(rp.vertAlign)) {
      drawSize = size * 0.65f;
      rise = -size * 0.14f;
    }
    final Layout.Style st = style(rp, link);
    final float spacing = nz(rp.spacing);
    for (int i = 0; i < text.length();) {
      final int cp = text.codePointAt(i);
      i += Character.charCount(cp);
      Item it = new Item(CHAR);
      it.cp = cp;
      it.outCp = cp == 0x2011 ? '-' : cp == 0x00A0 ? ' ' : cp;
      FontResolver.Face face = fonts.face(rp, it.outCp);
      it.face = face;
      it.size = size;
      it.drawSize = drawSize;
      it.rise = rise;
      it.natural = fonts.width(face.font(), it.outCp) * drawSize;
      // Word advances emboldened text (bold without a bold face) by 2% of the size more.
      it.adv = it.natural + spacing + (TextRules.eastAsian(cp) ? env.charDelta : 0f)
          + (face.syntheticBold() ? drawSize * SYNTHETIC_BOLD_ADVANCE : 0f);
      it.style = st;
      it.rpr = rp;
      items.add(it);
    }
  }

  /**
   * Applies the automatic spacing of East Asian typography: a quarter-em space between East
   * Asian and Latin text / digits, and the compression of consecutive full-width punctuation.
   */
  private void adjustSpacing(List<Item> items, ParaProps pp) {
    final boolean latin = !Boolean.FALSE.equals(pp.autoSpaceDe);
    final boolean digits = !Boolean.FALSE.equals(pp.autoSpaceDn);
    for (int i = 0; i + 1 < items.size(); i++) {
      Item a = items.get(i);
      Item b = items.get(i + 1);
      if (a.kind != CHAR || b.kind != CHAR) {
        continue;
      }
      boolean eaA = TextRules.eastAsianLetter(a.cp);
      boolean eaB = TextRules.eastAsianLetter(b.cp);
      if (eaA != eaB) {
        int other = eaA ? b.cp : a.cp;
        if ((latin && TextRules.latinLetter(other)) || (digits && TextRules.digit(other))) {
          a.autoSpace = (eaA ? a.size : b.size) * 0.25f;
          a.adv += a.autoSpace;
        }
      }
      if (settings.compressPunctuation && a.natural >= a.drawSize * 0.99f) {
        boolean openA = TextRules.fullWidthOpening(a.cp);
        boolean closeA = TextRules.fullWidthClosing(a.cp);
        boolean openB = TextRules.fullWidthOpening(b.cp);
        boolean closeB = TextRules.fullWidthClosing(b.cp);
        if ((closeA && (openB || closeB)) || (openA && openB)) {
          a.adv -= a.drawSize / 2;
          a.compressed = true;
        }
      }
    }
  }

  /** Returns whether the item is full-width punctuation which can still be compressed. */
  private static boolean compressible(Item it) {
    return it.kind == CHAR && !it.compressed && it.natural >= it.drawSize * 0.99f
        && (TextRules.fullWidthOpening(it.cp) || TextRules.fullWidthClosing(it.cp));
  }

  /**
   * Returns how much the space after (or before) the item can be reduced to shorten a line:
   * the blank half of full-width punctuation, the automatic space after the item, or a part
   * of a space between words.
   */
  private float shrinkable(Item it) {
    if (it.autoSpace > 0f) {
      return it.autoSpace;
    }
    if (it.space()) {
      return it.natural * SPACE_SHRINK;
    }
    return settings.compressPunctuation && compressible(it) ? it.drawSize / 2 : 0f;
  }

  /**
   * Returns how far the character {@code next} may go out of the line, since the line can be
   * shortened that much by compressing its punctuation (each by half a character at most)
   * and the automatic spaces between East Asian and Latin text.
   *
   * <p>Word compresses by up to half a character in total to put one more character on a
   * line. Only when the line cannot break before {@code next} by the kinsoku rules, it
   * compresses as much as possible to avoid moving the preceding character to the next
   * line.</p>
   */
  private float squeezable(List<Item> items, int from, int to, Item next, Item previous) {
    final boolean forced = TextRules.lineStartProhibited(next.cp)
        || (previous.kind == CHAR && TextRules.lineEndProhibited(previous.cp));
    float total = 0f;
    for (int j = from; j < to; j++) {
      Item it = items.get(j);
      if (it.space()) {
        continue;
      }
      float amount = shrinkable(it);
      total += amount;
    }
    return forced ? total : Math.min(total, next.drawSize / 2 + FIT_TOLERANCE);
  }

  /** Returns the index after the last item of the word which contains the item. */
  private static int wordEnd(List<Item> items, int index, boolean wordWrap) {
    int k = index + 1;
    while (k < items.size() && items.get(k).kind == CHAR && !items.get(k).space()
        && !breakable(items.get(k - 1), items.get(k), wordWrap)) {
      k++;
    }
    return k;
  }

  /**
   * Returns whether Word puts the word on the line by shrinking the spaces between words.
   *
   * <p>It does when each space shrinks by a quarter of its width at most, and that is at most
   * half of what each space would have to be expanded by if the word went to the next line
   * (measured on PDF files exported by Word).</p>
   *
   * @param lineStart the first item of the line
   * @param wordStart the first item of the word
   * @param index the item of the word which does not fit
   * @param wordEnd the index after the last item of the word
   * @param xs the positions of the items of the line up to {@code index}
   * @param x the position of the item at {@code index}
   */
  private static boolean fitsByShrinkingSpaces(List<Item> items, int lineStart, int wordStart,
      int index, int wordEnd, float[] xs, float x, float avail) {
    float end = x;
    for (int j = index; j < wordEnd; j++) {
      end += items.get(j).adv;
    }
    int beforeSpaces = wordStart;
    while (beforeSpaces > lineStart && items.get(beforeSpaces - 1).space()) {
      beforeSpaces--;
    }
    int spaces = 0;
    int spacesWithoutWord = 0;
    for (int j = lineStart; j < wordStart; j++) {
      if (items.get(j).space()) {
        spaces++;
        if (j < beforeSpaces) {
          spacesWithoutWord++;
        }
      }
    }
    if (spaces == 0) {
      return false;
    }
    float shrink = (end - avail) / spaces;
    float expand = spacesWithoutWord == 0 ? Float.MAX_VALUE
        : (avail - xs[beforeSpaces - lineStart]) / spacesWithoutWord;
    return shrink <= items.get(wordStart - 1).natural * SPACE_SHRINK && shrink <= expand / 2;
  }

  private static boolean breakable(Item before, Item after, boolean wordWrap) {
    if (before.kind != CHAR || after.kind != CHAR) {
      return true;
    }
    return TextRules.canBreak(before.cp, after.cp, wordWrap);
  }

  /** Returns whether a justified line is stretched between the two adjacent items. */
  private static boolean gapAfter(Item a, Item b) {
    if (a.kind != CHAR || b.kind != CHAR) {
      return false;
    }
    if (a.space()) {
      return true;
    }
    return !b.space() && (TextRules.eastAsian(a.cp) || TextRules.eastAsian(b.cp));
  }

  private float tabWidth(List<Item> items, int index, ParaProps pp, float absX, float left,
      float first, float right, float areaWidth) {
    Item tab = items.get(index);
    TabTarget target;
    if ("right".equals(tab.absoluteTab)) {
      target = new TabTarget(areaWidth - right, "right", null);
    } else if ("center".equals(tab.absoluteTab)) {
      target = new TabTarget(areaWidth / 2, "center", null);
    } else {
      target = nextTabStop(pp, absX, left, first);
    }
    tab.leader = target.leader();
    float width = target.pos() - absX;
    if (!"left".equals(target.align())) {
      float following = 0f;
      for (int j = index + 1; j < items.size(); j++) {
        Item it = items.get(j);
        if (it.kind != CHAR && it.kind != IMAGE) {
          break;
        }
        following += it.adv;
      }
      width -= "center".equals(target.align()) ? following / 2 : following;
    }
    return Math.max(0f, width);
  }

  private TabTarget nextTabStop(ParaProps pp, float absX, float left, float first) {
    TabTarget best = null;
    for (ParaProps.TabStop t : pp.tabs) {
      if (t.pos() > absX + EPS && !"bar".equals(t.align())) {
        best = new TabTarget(t.pos(), t.align(), t.leader());
        break;
      }
    }
    // A hanging indent works as a tab stop at the left indent.
    if (first < 0 && left > absX + EPS && (best == null || left < best.pos())) {
      best = new TabTarget(left, "left", null);
    }
    if (best != null) {
      return best;
    }
    float interval = settings.defaultTab > 0 ? settings.defaultTab : 36f;
    float pos = (float) (Math.floor((absX + EPS) / interval) + 1) * interval;
    return new TabTarget(pos, "left", null);
  }

  private void buildSegs(Layout.Line line, List<Item> items, int from, int to, float[] xs,
      float endX) {
    StringBuilder sb = new StringBuilder();
    Item head = null;
    float segX = 0f;
    float charSpacing = 0f;
    float expected = 0f;
    int count = 0;
    for (int j = from; j <= to; j++) {
      final Item it = j < to ? items.get(j) : null;
      final float x = j < to ? xs[j - from] : endX;
      boolean append = false;
      if (head != null && it != null && it.kind == CHAR && sameFormat(head, it)) {
        if (count == 1) {
          float cs = x - segX - head.natural;
          charSpacing = Math.abs(cs) < 0.001f ? 0f : cs;
          append = true;
        } else {
          append = Math.abs(x - expected) < 0.01f;
        }
      }
      if (!append && head != null) {
        FontResolver.Face face = head.face;
        line.segs.add(new Layout.Seg(segX, x - segX, sb.toString(),
            face == null ? null : face.font(), head.drawSize, count > 1 ? charSpacing : 0f,
            face != null && face.syntheticBold(), head.rise, head.style, null));
        sb.setLength(0);
        head = null;
        count = 0;
      }
      if (it == null) {
        break;
      }
      if (it.kind == CHAR) {
        if (head == null) {
          head = it;
          segX = x;
          charSpacing = 0f;
        }
        sb.appendCodePoint(it.outCp);
        count++;
        expected = x + it.natural + charSpacing;
      } else if (it.kind == TAB) {
        line.segs.add(new Layout.Seg(x, it.adv, "", null, it.size, 0f, false, 0f, it.style,
            null));
        addLeader(line, it, x);
      } else if (it.kind == IMAGE) {
        line.segs.add(new Layout.Seg(x, it.adv, "", null, 0f, 0f, false, 0f, PLAIN, it.image));
      }
    }
  }

  private static boolean sameFormat(Item a, Item b) {
    FontResolver.Face fa = a.face;
    FontResolver.Face fb = b.face;
    return fa != null && fb != null && fa.font().equals(fb.font())
        && fa.syntheticBold() == fb.syntheticBold() && Float.compare(a.drawSize, b.drawSize) == 0
        && Float.compare(a.rise, b.rise) == 0 && a.style.equals(b.style);
  }

  private void addLeader(Layout.Line line, Item tab, float x) {
    String leader = tab.leader;
    if (leader == null || tab.adv <= 0) {
      return;
    }
    int cp = switch (leader) {
      case "hyphen" -> '-';
      case "underscore", "heavy" -> '_';
      case "middleDot" -> 0x00B7;
      default -> '.';
    };
    FontResolver.Face face = fonts.face(tab.rpr, cp);
    float w = fonts.width(face.font(), cp) * tab.size;
    int count = (int) Math.floor((tab.adv - tab.size * 0.2f) / w);
    if (w <= 0 || count <= 0) {
      return;
    }
    String text = new String(Character.toChars(cp)).repeat(count);
    Layout.Style st = new Layout.Style(tab.style.color(), false, false, null, false,
        tab.style.link());
    line.segs.add(new Layout.Seg(x + tab.adv - count * w - tab.size * 0.1f, count * w, text,
        face.font(), tab.size, 0f, false, 0f, st, null));
  }

  /** Decides the height of the line and the position of its baseline. */
  private void measure(Layout.Line line, List<Item> items, int from, int to, Model.Para para,
      ParaProps pp, Layout.Env env) {
    float ascent = 0f;
    float descent = 0f;
    boolean any = false;
    for (int j = from; j < to; j++) {
      Item it = items.get(j);
      FontResolver.Face face = it.face;
      if (it.kind == TAB) {
        face = fonts.face(it.rpr, ' ');
      }
      if (face != null) {
        ascent = Math.max(ascent, face.ascent() * it.size);
        descent = Math.max(descent, face.descent() * it.size);
        any = true;
      }
      Model.Image image = it.image;
      if (image != null) {
        ascent = Math.max(ascent, image.height());
        any = true;
      }
    }
    if (!any) {
      FontResolver.Face face = fonts.face(para.markProps, ' ');
      float size = para.markProps.sizeOrDefault();
      ascent = face.ascent() * size;
      descent = face.descent() * size;
    }
    // Word works in twips (1/20 point): the height of the text is rounded to them.
    final float natural = Math.round((ascent + descent) * 20f) / 20f;
    final Float lineValue = pp.line;
    final String rule = pp.lineRule;
    if ("exact".equals(rule) && lineValue != null) {
      line.lineHeight = lineValue / 20f;
      line.baseline = line.lineHeight * 0.8f;
      return;
    }
    float height = natural;
    float baseline = ascent;
    final boolean grid = env.linePitch > 0 && !Boolean.FALSE.equals(pp.snapToGrid);
    if (grid) {
      // Lines snap to the document grid: the text is centered in a multiple of the pitch.
      int count = Math.max(1, (int) Math.ceil(natural / env.linePitch - 0.001f));
      height = count * env.linePitch;
      baseline = (height - natural) / 2 + ascent;
    }
    if (lineValue != null && "atLeast".equals(rule)) {
      // The additional space is put above the text.
      float target = Math.max(height, lineValue / 20f);
      baseline += target - height;
      height = target;
    } else if (lineValue != null) {
      // A multiple of the line height. On the grid the text stays centered in the line.
      // Without the grid the additional space is put below the text.
      float target = height * lineValue / 240f;
      if (grid) {
        baseline += (target - height) / 2;
      }
      height = target;
    }
    line.lineHeight = height;
    line.baseline = baseline;
    // The space below the text (of a multiple line height, or of the grid) may go out of
    // the bottom of a page: a line fits when its text does.
    line.slack = Math.max(0f, height - baseline - descent);
  }
}
