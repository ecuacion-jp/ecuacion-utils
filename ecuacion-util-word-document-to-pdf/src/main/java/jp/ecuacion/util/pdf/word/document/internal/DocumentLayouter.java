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

import java.util.IdentityHashMap;
import java.util.List;
import org.jspecify.annotations.Nullable;

/** Flows the content of a document into pages. */
final class DocumentLayouter {

  private static final int MAX_PASSES = 4;
  private static final float EPS = 0.01f;

  private final Model.Document doc;
  private final ParagraphLayouter paragraphs;
  private final BlockLayouter blocks;

  // state of a pass
  private Layout.PassResult result = new Layout.PassResult();
  private Layout.@Nullable PassResult previous;
  private Layout.@Nullable PageCanvas page;
  private float pendingAfter;
  private Model.@Nullable Para previousPara;
  /**
   * Whether "space before" is ignored at the top of the current page: Word ignores it on
   * every page except the first page of a section.
   */
  private boolean startedByOverflow;
  private IdentityHashMap<Model.Para, List<Layout.Line>> lineCache = new IdentityHashMap<>();
  private IdentityHashMap<Model.Table, List<BlockLayouter.RowLayout>> tableCache =
      new IdentityHashMap<>();

  DocumentLayouter(Model.Document doc, FontResolver fonts) {
    this.doc = doc;
    this.paragraphs = new ParagraphLayouter(fonts, doc.settings);
    this.blocks = new BlockLayouter(paragraphs, doc.settings);
  }

  /**
   * Lays out the document. Since page numbers shown in the body (e.g. in a table of contents)
   * are known only after the layout, it is repeated until they do not change.
   */
  Layout.PassResult layout() {
    boolean dynamic = false;
    for (Model.Section s : doc.sections) {
      dynamic |= hasDynamicField(s.blocks);
    }
    Layout.PassResult last = runPass(null);
    for (int pass = 1; dynamic && pass < MAX_PASSES; pass++) {
      Layout.PassResult next = runPass(last);
      boolean stable = next.bookmarkPages.equals(last.bookmarkPages)
          && next.pages.size() == last.pages.size();
      last = next;
      if (stable && pass > 1) {
        break;
      }
    }
    return last;
  }

  private static boolean hasDynamicField(List<Model.Block> list) {
    for (Model.Block b : list) {
      if (b instanceof Model.Para p) {
        for (Model.Inline inline : p.inlines) {
          if (inline instanceof Model.Field) {
            return true;
          }
        }
      } else if (b instanceof Model.Table t) {
        for (Model.Row row : t.rows) {
          for (Model.Cell cell : row.cells) {
            if (hasDynamicField(cell.blocks)) {
              return true;
            }
          }
        }
      }
    }
    return false;
  }

  private Layout.PassResult runPass(Layout.@Nullable PassResult prev) {
    result = new Layout.PassResult();
    previous = prev;
    page = null;
    pendingAfter = 0f;
    previousPara = null;
    lineCache = new IdentityHashMap<>();
    tableCache = new IdentityHashMap<>();

    for (Model.Section section : doc.sections) {
      Layout.PageCanvas current = page;
      if (current == null || !"continuous".equals(section.type)
          || differs(current.section, section)) {
        startPage(section, true);
      }
      for (int i = 0; i < section.blocks.size(); i++) {
        Model.Block b = section.blocks.get(i);
        if (b instanceof Model.Para p) {
          placePara(section, i, p);
        } else if (b instanceof Model.Table t) {
          placeTable(section, t);
        }
      }
    }

    for (Layout.PageCanvas c : result.pages) {
      Model.Section s = c.section;
      Layout.Env env = headerEnv(c, s);
      env.numPages = Integer.toString(result.pages.size());
      env.sectionPages = Integer.toString(sectionPageCount(result, s));
      env.bookmarkPages = result.bookmarkPages;
      float width = s.pageWidth - s.marginLeft - s.marginRight;
      float y = s.headerDistance;
      for (Layout.FlowItem item : blocks.flow(c.header, width, env)) {
        item.draw(c, s.marginLeft, y);
        y += item.height();
      }
      List<Layout.FlowItem> footer = blocks.flow(c.footer, width, env);
      y = s.pageHeight - s.footerDistance - BlockLayouter.height(footer);
      for (Layout.FlowItem item : footer) {
        item.draw(c, s.marginLeft, y);
        y += item.height();
      }
    }
    return result;
  }

  /** Returns the number of the pages which belong to the section (start on it). */
  private static int sectionPageCount(Layout.PassResult result, Model.Section section) {
    int count = 0;
    for (Layout.PageCanvas c : result.pages) {
      if (c.section == section) {
        count++;
      }
    }
    return count;
  }

  private static boolean differs(Model.Section a, Model.Section b) {
    return Math.abs(a.pageWidth - b.pageWidth) > EPS
        || Math.abs(a.pageHeight - b.pageHeight) > EPS
        || Math.abs(a.marginLeft - b.marginLeft) > EPS
        || Math.abs(a.marginRight - b.marginRight) > EPS;
  }

  private Layout.Env env(Layout.PageCanvas c, Model.Section section) {
    Layout.Env env = new Layout.Env();
    env.pageNumber = c.number;
    env.pageText = c.numberText;
    Layout.PassResult prev = previous;
    if (prev != null) {
      env.numPages = Integer.toString(prev.pages.size());
      env.sectionPages = Integer.toString(sectionPageCount(prev, section));
      env.bookmarkPages = prev.bookmarkPages;
      env.fieldPages = prev.fieldPages;
    }
    env.linePitch = section.linePitch;
    env.charDelta = section.charDelta;
    env.charUnit = doc.defaultRun.sizeOrDefault() + section.charDelta;
    return env;
  }

  /**
   * Returns the values for a header or a footer of the page: lines do not snap to the grid,
   * and a PAGE field is always the number of the page (the same header is on many pages).
   */
  private Layout.Env headerEnv(Layout.PageCanvas c, Model.Section section) {
    Layout.Env env = env(c, section);
    env.linePitch = 0f;
    env.fieldPages = new IdentityHashMap<>();
    return env;
  }

  private Layout.PageCanvas page() {
    Layout.PageCanvas c = page;
    if (c == null) {
      throw new IllegalStateException("No page has been started.");
    }
    return c;
  }

  private Layout.PageCanvas startPage(Model.Section section, boolean firstOfSection) {
    Layout.PageCanvas prev = page;
    Layout.PageCanvas c = new Layout.PageCanvas(section, result.pages.size(), result);
    Integer start = section.pageNumberStart;
    c.number = firstOfSection && start != null ? start : prev == null ? 1 : prev.number + 1;
    c.numberText = NumberFormats.format(c.number, section.pageNumberFormat);
    String kind = firstOfSection && section.titlePage ? "first"
        : doc.settings.evenAndOddHeaders && c.number % 2 == 0 ? "even" : "default";
    c.header = section.headers.getOrDefault(kind, List.of());
    c.footer = section.footers.getOrDefault(kind, List.of());
    float width = section.pageWidth - section.marginLeft - section.marginRight;
    Layout.Env env = headerEnv(c, section);
    float headerHeight = BlockLayouter.height(blocks.flow(c.header, width, env));
    float footerHeight = BlockLayouter.height(blocks.flow(c.footer, width, env));
    c.bodyTop = Math.max(section.marginTop, section.headerDistance + headerHeight);
    c.bodyBottom = section.pageHeight
        - Math.max(section.marginBottom, section.footerDistance + footerHeight);
    c.cursorY = c.bodyTop;
    result.pages.add(c);
    page = c;
    pendingAfter = 0f;
    startedByOverflow = !firstOfSection;
    return c;
  }

  private List<Layout.Line> lines(Model.Section section, Model.Para p) {
    List<Layout.Line> lines = lineCache.get(p);
    if (lines == null) {
      float width = section.pageWidth - section.marginLeft - section.marginRight;
      lines = paragraphs.layout(p, width, env(page(), section));
      lineCache.put(p, lines);
    }
    return lines;
  }

  private List<BlockLayouter.RowLayout> rows(Model.Section section, Model.Table t) {
    List<BlockLayouter.RowLayout> rows = tableCache.get(t);
    if (rows == null) {
      float width = section.pageWidth - section.marginLeft - section.marginRight;
      rows = blocks.table(t, width, env(page(), section));
      tableCache.put(t, rows);
    }
    return rows;
  }

  // ---------------------------------------------------------------------------------------
  // paragraphs
  // ---------------------------------------------------------------------------------------

  private void placePara(Model.Section section, int index, Model.Para p) {
    if (Boolean.TRUE.equals(p.props.pageBreakBefore) && !page().empty) {
      startPage(section, false);
    }
    final List<Layout.Line> lines = lines(section, p);
    final int n = lines.size();
    final Layout.Line tail = lines.get(n - 1);
    final float after = tail.spaceAfter;
    tail.spaceAfter = 0f;

    float gap = pendingAfter;
    Model.Para prev = previousPara;
    // "Don't add space between paragraphs of the same style": when either of two adjacent
    // paragraphs of the same style has it, there is no space between them.
    if (prev != null && java.util.Objects.equals(prev.props.styleId, p.props.styleId)
        && (Boolean.TRUE.equals(prev.props.contextualSpacing)
            || Boolean.TRUE.equals(p.props.contextualSpacing))) {
      gap = 0f;
      lines.get(0).spaceBefore = 0f;
    }
    if (!page().empty && Boolean.TRUE.equals(p.props.keepNext)
        && keepNeedsNewPage(section, index, lines, gap)) {
      startPage(section, false);
    }

    final boolean widowControl = !Boolean.FALSE.equals(p.props.widowControl);
    final boolean keepLines = Boolean.TRUE.equals(p.props.keepLines);
    int idx = 0;
    while (idx < n) {
      Layout.PageCanvas c = page();
      if (c.empty) {
        gap = 0f;
        if (startedByOverflow && idx == 0) {
          // Word ignores "space before" at the top of a page.
          lines.get(0).spaceBefore = 0f;
        }
      }
      int chunkEnd = idx;
      while (chunkEnd < n - 1 && !lines.get(chunkEnd).pageBreakAfter) {
        chunkEnd++;
      }
      final int rest = chunkEnd - idx + 1;
      final float remaining = c.bodyBottom - c.cursorY - gap;
      int fit = 0;
      float h = 0f;
      while (fit < rest && h + lines.get(idx + fit).height() - lines.get(idx + fit).slack
          <= remaining + EPS) {
        h += lines.get(idx + fit).height();
        fit++;
      }
      int take = fit;
      if (fit < rest) {
        if (keepLines && idx == 0) {
          take = 0;
        } else if (widowControl) {
          boolean paragraphEnd = chunkEnd == n - 1;
          if (paragraphEnd && rest - take == 1) {
            take--;
          }
          if (idx == 0 && take == 1) {
            take = 0;
          }
        }
        if (take <= 0 && c.empty) {
          take = Math.max(1, fit);
        }
      }
      if (take <= 0) {
        startPage(section, false);
        gap = 0f;
        continue;
      }
      c.cursorY += gap;
      gap = 0f;
      for (int k = idx; k < idx + take; k++) {
        Layout.Line line = lines.get(k);
        line.draw(c, section.marginLeft, c.cursorY);
        c.cursorY += line.height();
      }
      c.empty = false;
      idx += take;
      if (take < rest) {
        startPage(section, false);
      } else if (lines.get(chunkEnd).pageBreakAfter) {
        startPage(section, false);
      }
    }
    pendingAfter = page().empty ? 0f : after;
    previousPara = p;
  }

  /** Returns whether the paragraph must move to the next page to stay with the next one. */
  private boolean keepNeedsNewPage(Model.Section section, int index, List<Layout.Line> lines,
      float gap) {
    float need = gap + BlockLayouter.height(lines);
    List<Model.Block> list = section.blocks;
    for (int j = index + 1, guard = 0; j < list.size() && guard < 8; j++, guard++) {
      Model.Block next = list.get(j);
      if (next instanceof Model.Para q) {
        List<Layout.Line> ql = lines(section, q);
        if (Boolean.TRUE.equals(q.props.keepNext)) {
          need += BlockLayouter.height(ql);
          continue;
        }
        boolean widowControl = !Boolean.FALSE.equals(q.props.widowControl);
        int count = !widowControl ? 1 : ql.size() == 3 ? 3 : Math.min(2, ql.size());
        for (int k = 0; k < count; k++) {
          need += ql.get(k).height();
        }
      } else if (next instanceof Model.Table t) {
        List<BlockLayouter.RowLayout> rows = rows(section, t);
        if (!rows.isEmpty()) {
          need += rows.get(0).height;
        }
      }
      break;
    }
    Layout.PageCanvas c = page();
    return need > c.bodyBottom - c.cursorY + EPS && need <= c.bodyBottom - c.bodyTop;
  }

  // ---------------------------------------------------------------------------------------
  // tables
  // ---------------------------------------------------------------------------------------

  private void placeTable(Model.Section section, Model.Table t) {
    final List<BlockLayouter.RowLayout> rows = rows(section, t);
    int headerCount = 0;
    while (headerCount < rows.size() && rows.get(headerCount).row.header) {
      headerCount++;
    }
    if (headerCount == rows.size()) {
      headerCount = 0;
    }
    if (!page().empty) {
      page().cursorY += pendingAfter;
    }
    pendingAfter = 0f;
    previousPara = null;
    float pageStart = page().empty ? page().cursorY : Float.NaN;

    for (int r = 0; r < rows.size(); r++) {
      BlockLayouter.RowLayout row = rows.get(r);
      final float x = section.marginLeft + row.tableX;
      int[] from = new int[row.cells.size()];
      boolean split = false;
      while (true) {
        Layout.PageCanvas c = page();
        float remaining = c.bodyBottom - c.cursorY;
        float h = row.restHeight(from);
        boolean atTop = Math.abs(c.cursorY - pageStart) < EPS;
        if (h <= remaining + EPS) {
          row.drawPart(c, x, c.cursorY, h, split ? from : null, null);
          c.cursorY += h;
          c.empty = false;
          break;
        }
        if (!row.row.cantSplit && r >= headerCount) {
          int[] to = row.fit(from, remaining);
          if (to != null) {
            row.drawPart(c, x, c.cursorY, remaining, from, to);
            c.empty = false;
            from = to;
            split = true;
            pageStart = continueOnNextPage(section, rows, headerCount, r);
            continue;
          }
        }
        if (atTop) {
          // Even a whole page is not enough: output it as it is.
          row.drawPart(c, x, c.cursorY, h, split ? from : null, null);
          c.cursorY += h;
          c.empty = false;
          break;
        }
        pageStart = continueOnNextPage(section, rows, headerCount, r);
      }
    }
  }

  /** Starts a new page for a table, repeating its header rows. Returns the y after them. */
  private float continueOnNextPage(Model.Section section, List<BlockLayouter.RowLayout> rows,
      int headerCount, int currentRow) {
    Layout.PageCanvas c = startPage(section, false);
    if (currentRow >= headerCount) {
      float headerHeight = 0f;
      for (int i = 0; i < headerCount; i++) {
        headerHeight += rows.get(i).height;
      }
      if (headerHeight < (c.bodyBottom - c.bodyTop) / 2) {
        for (int i = 0; i < headerCount; i++) {
          BlockLayouter.RowLayout header = rows.get(i);
          header.draw(c, section.marginLeft, c.cursorY);
          c.cursorY += header.height;
        }
      }
    }
    return c.cursorY;
  }
}
