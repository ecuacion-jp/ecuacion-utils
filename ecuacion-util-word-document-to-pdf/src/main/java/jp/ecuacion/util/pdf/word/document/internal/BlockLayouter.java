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
import java.util.Map;
import org.jspecify.annotations.Nullable;

/** Lays out blocks (paragraphs and tables) into items which flow vertically. */
final class BlockLayouter {

  /** The layout of a table cell. Positions are relative to the left edge of the table. */
  static final class CellLayout {
    final Model.Cell cell;
    float offsetX;
    float width;
    final float[] margins = new float[4];
    List<Layout.FlowItem> items = new ArrayList<>();
    float contentHeight;
    int rowSpan = 1;
    /** The total height of the rows a vertically merged cell covers. */
    float spanHeight;
    @Nullable Border top;
    @Nullable Border bottom;
    @Nullable Border left;
    @Nullable Border right;

    CellLayout(Model.Cell cell) {
      this.cell = cell;
    }

    float paddedHeight(int from) {
      float h = margins[1] + margins[3];
      for (int i = from; i < items.size(); i++) {
        h += items.get(i).height();
      }
      return h;
    }
  }

  /** The layout of a table row. */
  static final class RowLayout implements Layout.FlowItem {
    final Model.Row row;
    final List<CellLayout> cells = new ArrayList<>();
    float height;
    /** The height the top border takes, which a minimum row height does not include. */
    float borderRoom;
    /** The left edge of the table from the left edge of the area. */
    float tableX;

    RowLayout(Model.Row row) {
      this.row = row;
    }

    @Override
    public float height() {
      return height;
    }

    @Override
    public void draw(Layout.PageCanvas canvas, float x, float y) {
      drawPart(canvas, x + tableX, y, height, null, null);
    }

    /** Returns the height of the rest of the row, when split from the item indexes. */
    float restHeight(int[] from) {
      boolean started = false;
      float h = 0f;
      for (int c = 0; c < cells.size(); c++) {
        started |= from[c] > 0;
        CellLayout cl = cells.get(c);
        if (cl.rowSpan == 1 && cl.cell.vmerge != 2) {
          h = Math.max(h, cl.paddedHeight(from[c]));
        }
      }
      return started ? h : height;
    }

    /**
     * Returns the item indexes up to which each cell fits into the height, or {@code null}
     * when some cell cannot put even one more item there.
     */
    int @Nullable [] fit(int[] from, float available) {
      int[] to = new int[cells.size()];
      for (int c = 0; c < cells.size(); c++) {
        CellLayout cl = cells.get(c);
        int size = cl.items.size();
        if (cl.rowSpan > 1 || from[c] >= size) {
          to[c] = size;
          continue;
        }
        float h = cl.margins[1] + cl.margins[3];
        int k = from[c];
        while (k < size && h + cl.items.get(k).height() <= available + 0.01f) {
          h += cl.items.get(k).height();
          k++;
        }
        if (k == from[c]) {
          return null;
        }
        to[c] = k;
      }
      return to;
    }

    /**
     * Draws the row, or a part of it when the row is split across pages.
     *
     * @param x the left edge of the table
     * @param from the first item index of each cell, or {@code null} for all the items
     * @param to the item index each cell is drawn up to, or {@code null} for all the items
     */
    void drawPart(Layout.PageCanvas canvas, float x, float y, float partHeight,
        int @Nullable [] from, int @Nullable [] to) {
      final boolean whole = from == null && to == null;
      for (CellLayout cl : cells) {
        Integer shade = cl.cell.shade;
        if (shade != null && shade >= 0 && cl.cell.vmerge != 2) {
          float h = whole && cl.rowSpan > 1 ? cl.spanHeight : partHeight;
          float cx = x + cl.offsetX;
          canvas.add(p -> p.fillRect(cx, y, cl.width, h, shade));
        }
      }
      for (int c = 0; c < cells.size(); c++) {
        CellLayout cl = cells.get(c);
        int start = from == null ? 0 : from[c];
        int end = to == null ? cl.items.size() : to[c];
        float cy = y + cl.margins[1];
        if (whole) {
          float space = (cl.rowSpan > 1 ? cl.spanHeight : partHeight) - cl.margins[1]
              - cl.margins[3] - cl.contentHeight;
          if (space > 0 && "center".equals(cl.cell.valign)) {
            cy += space / 2;
          } else if (space > 0 && "bottom".equals(cl.cell.valign)) {
            cy += space;
          }
        }
        for (int i = start; i < end; i++) {
          Layout.FlowItem item = cl.items.get(i);
          item.draw(canvas, x + cl.offsetX + cl.margins[0], cy);
          cy += item.height();
        }
      }
      for (CellLayout cl : cells) {
        float x1 = x + cl.offsetX;
        float x2 = x1 + cl.width;
        float y2 = y + partHeight;
        border(canvas, cl.top, x1, y, x2, y);
        border(canvas, cl.bottom, x1, y2, x2, y2);
        border(canvas, cl.left, x1, y, x1, y2);
        border(canvas, cl.right, x2, y, x2, y2);
      }
    }

    private static void border(Layout.PageCanvas canvas, @Nullable Border b, float x1, float y1,
        float x2, float y2) {
      if (b != null && b.visible()) {
        canvas.add(p -> p.line(x1, y1, x2, y2, b.width(), b.color(), b.style()));
      }
    }
  }

  private final ParagraphLayouter paragraphs;
  private final Model.Settings settings;

  BlockLayouter(ParagraphLayouter paragraphs, Model.Settings settings) {
    this.paragraphs = paragraphs;
    this.settings = settings;
  }

  /** Lays out blocks in an area of the width (a table cell, a header or a footer). */
  List<Layout.FlowItem> flow(List<Model.Block> blocks, float width, Layout.Env env) {
    List<Layout.FlowItem> items = new ArrayList<>();
    for (Model.Block b : blocks) {
      if (b instanceof Model.Para p) {
        items.addAll(paragraphs.layout(p, width, env));
      } else if (b instanceof Model.Table t) {
        items.addAll(table(t, width, env));
      }
    }
    return items;
  }

  static float height(List<? extends Layout.FlowItem> items) {
    float h = 0f;
    for (Layout.FlowItem item : items) {
      h += item.height();
    }
    return h;
  }

  /** Lays out a table. */
  List<RowLayout> table(Model.Table t, float available, Layout.Env outerEnv) {
    final Layout.Env env = outerEnv.withGrid(settings.snapToGridInCell);
    int columns = t.grid.length;
    for (Model.Row row : t.rows) {
      int count = row.gridBefore;
      for (Model.Cell cell : row.cells) {
        count += cell.span;
      }
      columns = Math.max(columns, count);
    }
    float[] widths = new float[columns];
    float total = 0f;
    for (int i = 0; i < columns; i++) {
      widths[i] = i < t.grid.length ? t.grid[i] : 0f;
      total += widths[i];
    }
    if (total <= 0f) {
      for (int i = 0; i < columns; i++) {
        widths[i] = available / Math.max(1, columns);
      }
      total = available;
    }
    Float ratio = t.widthRatio;
    if (ratio != null && ratio > 0f) {
      float scale = available * ratio / total;
      for (int i = 0; i < columns; i++) {
        widths[i] *= scale;
      }
      total *= scale;
    }
    if (!t.fixedLayout) {
      fitToContent(t, widths, env);
    }
    float[] colX = new float[columns + 1];
    for (int i = 0; i < columns; i++) {
      colX[i + 1] = colX[i] + widths[i];
    }
    float tableX;
    if ("center".equals(t.jc)) {
      tableX = (available - total) / 2;
    } else if ("right".equals(t.jc)) {
      tableX = available - total;
    } else {
      // Before Word 2013 the text in the cells, not the edge of the table, is aligned.
      tableX = t.indent - (settings.compatibilityMode < 15 ? t.cellMargins[0] : 0f);
    }

    List<RowLayout> rows = new ArrayList<>();
    List<Map<Integer, CellLayout>> byColumn = new ArrayList<>();
    for (int r = 0; r < t.rows.size(); r++) {
      Model.Row row = t.rows.get(r);
      RowLayout rl = new RowLayout(row);
      rl.tableX = tableX;
      Map<Integer, CellLayout> map = new HashMap<>();
      float borderRoom = 0f;
      int col = Math.min(row.gridBefore, columns);
      for (int c = 0; c < row.cells.size(); c++) {
        Model.Cell cell = row.cells.get(c);
        CellLayout cl = new CellLayout(cell);
        int endCol = Math.min(col + cell.span, columns);
        cl.offsetX = colX[col];
        cl.width = colX[endCol] - colX[col];
        for (int i = 0; i < 4; i++) {
          Float m = cell.margins[i];
          cl.margins[i] = m != null ? m : t.cellMargins[i];
        }
        final boolean lastRow = r == t.rows.size() - 1;
        final boolean lastCell = c == row.cells.size() - 1;
        cl.top = border(cell, t, "top", r == 0 ? "top" : "insideH");
        cl.bottom = border(cell, t, "bottom", lastRow ? "bottom" : "insideH");
        cl.left = border(cell, t, "left", c == 0 ? "left" : "insideV");
        cl.right = border(cell, t, "right", lastCell ? "right" : "insideV");
        // Borders take room: the top border of each row and the bottom border of the table
        // add to the height, and the content moves right by half of the left border.
        Border topBorder = cl.top;
        if (topBorder != null && topBorder.visible()) {
          cl.margins[1] += topBorder.room();
          borderRoom = Math.max(borderRoom, topBorder.room());
        }
        Border bottomBorder = cl.bottom;
        if (lastRow && bottomBorder != null && bottomBorder.visible()) {
          cl.margins[3] += bottomBorder.room();
        }
        Border leftBorder = cl.left;
        if (leftBorder != null && leftBorder.visible()) {
          cl.margins[0] += leftBorder.room() / 2;
        }
        if (cell.vmerge != 2) {
          cl.items = flow(cell.blocks, Math.max(1f, cl.width - cl.margins[0] - cl.margins[2]),
              env);
          cl.contentHeight = height(cl.items);
        }
        map.put(col, cl);
        rl.cells.add(cl);
        col = endCol;
      }
      // The top and bottom margins of a row are the largest ones of its cells.
      float top = 0f;
      float bottom = 0f;
      for (CellLayout cl : rl.cells) {
        top = Math.max(top, cl.margins[1]);
        bottom = Math.max(bottom, cl.margins[3]);
      }
      for (CellLayout cl : rl.cells) {
        cl.margins[1] = top;
        cl.margins[3] = bottom;
      }
      rl.borderRoom = borderRoom;
      rows.add(rl);
      byColumn.add(map);
    }

    // vertical merges
    for (int r = 0; r < rows.size(); r++) {
      for (Map.Entry<Integer, CellLayout> e : byColumn.get(r).entrySet()) {
        CellLayout cl = e.getValue();
        if (cl.cell.vmerge != 1) {
          continue;
        }
        CellLayout lastOfSpan = cl;
        for (int rr = r + 1; rr < rows.size(); rr++) {
          CellLayout below = byColumn.get(rr).get(e.getKey());
          if (below == null || below.cell.vmerge != 2) {
            break;
          }
          lastOfSpan.bottom = null;
          below.top = null;
          lastOfSpan = below;
          cl.rowSpan++;
        }
      }
    }
    for (RowLayout rl : rows) {
      float h = rl.row.height;
      if (!rl.row.exactHeight) {
        h = h > 0f ? h + rl.borderRoom : 0f;
        for (CellLayout cl : rl.cells) {
          if (cl.rowSpan == 1 && cl.cell.vmerge != 2) {
            h = Math.max(h, cl.paddedHeight(0));
          }
        }
      }
      rl.height = h;
    }
    for (int r = 0; r < rows.size(); r++) {
      for (CellLayout cl : rows.get(r).cells) {
        if (cl.rowSpan > 1) {
          int lastRow = Math.min(rows.size(), r + cl.rowSpan) - 1;
          float sum = 0f;
          for (int rr = r; rr <= lastRow; rr++) {
            sum += rows.get(rr).height;
          }
          float need = cl.paddedHeight(0);
          if (need > sum) {
            rows.get(lastRow).height += need - sum;
            sum = need;
          }
          cl.spanHeight = sum;
        }
      }
    }
    // The heights of the rows below may have grown after spanHeight was computed.
    for (int r = 0; r < rows.size(); r++) {
      for (CellLayout cl : rows.get(r).cells) {
        if (cl.rowSpan > 1) {
          float sum = 0f;
          for (int rr = r; rr < Math.min(rows.size(), r + cl.rowSpan); rr++) {
            sum += rows.get(rr).height;
          }
          cl.spanHeight = sum;
        }
      }
    }
    return rows;
  }

  /**
   * Widens the columns which are narrower than what cannot be broken in them (e.g. a long
   * word), taking the width from the columns which have room, as Word does for tables whose
   * layout is not fixed.
   */
  private void fitToContent(Model.Table t, float[] widths, Layout.Env env) {
    float[] min = new float[widths.length];
    for (Model.Row row : t.rows) {
      int col = row.gridBefore;
      for (Model.Cell cell : row.cells) {
        if (cell.span == 1 && col < widths.length && cell.vmerge != 2) {
          float content = 0f;
          for (Model.Block b : cell.blocks) {
            if (b instanceof Model.Para para) {
              content = Math.max(content, paragraphs.minContentWidth(para, env));
            }
          }
          Float left = cell.margins[0];
          Float right = cell.margins[2];
          float margins = (left != null ? left : t.cellMargins[0])
              + (right != null ? right : t.cellMargins[2]);
          Border vertical = t.borders.get(col == 0 ? "left" : "insideV");
          float border = vertical != null && vertical.visible() ? vertical.room() / 2 : 0f;
          min[col] = Math.max(min[col], content + margins + border);
        }
        col += cell.span;
      }
    }
    float need = 0f;
    float slack = 0f;
    for (int i = 0; i < widths.length; i++) {
      if (min[i] > widths[i]) {
        need += min[i] - widths[i];
      } else {
        slack += widths[i] - min[i];
      }
    }
    if (need <= 0.01f || slack <= 0f) {
      return;
    }
    float ratio = Math.min(1f, need / slack);
    for (int i = 0; i < widths.length; i++) {
      if (min[i] > widths[i]) {
        widths[i] += (min[i] - widths[i]) * Math.min(1f, slack / need);
      } else {
        widths[i] -= (widths[i] - min[i]) * ratio;
      }
    }
  }

  private static @Nullable Border border(Model.Cell cell, Model.Table t, String edge,
      String tableEdge) {
    Border b = cell.borders.get(edge);
    return b != null ? b : t.borders.get(tableEdge);
  }
}
