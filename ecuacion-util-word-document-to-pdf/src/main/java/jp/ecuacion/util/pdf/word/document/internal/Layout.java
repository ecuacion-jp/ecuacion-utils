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

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.jspecify.annotations.Nullable;

/** Holds the types shared by the layout classes. Lengths are in points, y grows downwards. */
final class Layout {

  private Layout() {}

  /** Something which occupies a height in a flow and can be drawn at a position. */
  interface FlowItem {
    float height();

    void draw(PageCanvas canvas, float x, float y);
  }

  /** A drawing operation recorded on a page. */
  interface Op {
    void paint(PdfPainter painter) throws IOException;
  }

  /** The values which depend on where a paragraph is laid out. */
  static final class Env {
    int pageNumber = 1;
    String pageText = "";
    String numPages = "";
    /** The number of the pages of the section (SECTIONPAGES). */
    String sectionPages = "";
    Map<String, String> bookmarkPages = new HashMap<>();
    /** PAGE field in the body -> the page it was on in the previous layout pass. */
    IdentityHashMap<Model.Field, PageCanvas> fieldPages = new IdentityHashMap<>();
    /** The line pitch of the document grid, or 0. */
    float linePitch;
    float charDelta;
    /** The width of a character, the unit of indents specified in characters. */
    float charUnit = 10.5f;

    /**
     * Returns a copy for table cells (where lines may not snap to the document grid), and for
     * headers and footers (where lines never snap to it).
     */
    Env withGrid(boolean snapToGrid) {
      if (snapToGrid || linePitch == 0f) {
        return this;
      }
      Env e = new Env();
      e.pageNumber = pageNumber;
      e.pageText = pageText;
      e.numPages = numPages;
      e.sectionPages = sectionPages;
      e.bookmarkPages = bookmarkPages;
      e.fieldPages = fieldPages;
      e.charDelta = charDelta;
      e.charUnit = charUnit;
      return e;
    }
  }

  /** The decoration of a piece of text. */
  record Style(int color, boolean underline, boolean strike, @Nullable Integer highlight,
      boolean italic, @Nullable String link) {
  }

  /** A piece of a line drawn with one font and style ({@code font} is null for a gap). */
  record Seg(float x, float width, String text, @Nullable PDType0Font font, float size,
      float charSpacing, boolean syntheticBold, float rise, Style style,
      Model.@Nullable Image image) {
  }

  /** A heading for the outline (bookmarks) of the PDF file. */
  record Outline(String title, int level, int page, float y) {
  }

  /** The data collected by a layout pass, used by the next pass and by the painter. */
  static final class PassResult {
    final List<PageCanvas> pages = new ArrayList<>();
    final Map<String, String> bookmarkPages = new HashMap<>();
    /** bookmark name -> {page index, y}. */
    final Map<String, float[]> bookmarkPositions = new HashMap<>();
    final List<Outline> outlines = new ArrayList<>();
    /** PAGE field -> the page its line was drawn on. */
    final IdentityHashMap<Model.Field, PageCanvas> fieldPages = new IdentityHashMap<>();
  }

  /** A page being laid out: records drawing operations. */
  static final class PageCanvas {
    final Model.Section section;
    final int index;
    final PassResult result;
    final List<Op> ops = new ArrayList<>();
    int number = 1;
    String numberText = "1";
    float bodyTop;
    float bodyBottom;
    /** The current position of the body flow. */
    float cursorY;
    boolean empty = true;
    List<Model.Block> header = List.of();
    List<Model.Block> footer = List.of();

    PageCanvas(Model.Section section, int index, PassResult result) {
      this.section = section;
      this.index = index;
      this.result = result;
    }

    void add(Op op) {
      ops.add(op);
    }

    void bookmark(String name, float y) {
      if (!result.bookmarkPositions.containsKey(name)) {
        result.bookmarkPositions.put(name, new float[] {index, y});
        result.bookmarkPages.put(name, numberText);
      }
    }

    void outline(String title, int level, float y) {
      result.outlines.add(new Outline(title, level, index, y));
    }
  }

  /** A line of a paragraph. */
  static final class Line implements FlowItem {
    final List<Seg> segs = new ArrayList<>();
    /** The start of the text from the left edge of the area. */
    float x0;
    float lineHeight;
    /** The baseline from the top of the line. */
    float baseline;
    float spaceBefore;
    float spaceAfter;
    /** The space below the text which may go out of the bottom of a page. */
    float slack;
    float padTop;
    float padBottom;
    @Nullable Integer background;
    float boxLeft;
    float boxRight;
    @Nullable Border borderTop;
    @Nullable Border borderBottom;
    @Nullable Border borderLeft;
    @Nullable Border borderRight;
    final List<String> bookmarks = new ArrayList<>();
    final List<Model.Image> floating = new ArrayList<>();
    final List<Model.Field> pageFields = new ArrayList<>();
    @Nullable String outlineTitle;
    int outlineLevel;
    boolean pageBreakAfter;

    @Override
    public float height() {
      return spaceBefore + padTop + lineHeight + padBottom + spaceAfter;
    }

    @Override
    public void draw(PageCanvas canvas, float x, float y) {
      final float top = y + spaceBefore;
      final float boxHeight = padTop + lineHeight + padBottom;
      final float lineTop = top + padTop;
      final float base = lineTop + baseline;
      final float left = x + boxLeft;
      final float width = boxRight - boxLeft;
      Integer bg = background;
      if (bg != null) {
        canvas.add(p -> p.fillRect(left, top, width, boxHeight, bg));
      }
      for (String name : bookmarks) {
        canvas.bookmark(name, top);
      }
      for (Model.Field field : pageFields) {
        canvas.result.fieldPages.put(field, canvas);
      }
      String title = outlineTitle;
      if (title != null) {
        canvas.outline(title, outlineLevel, top);
      }
      for (Seg seg : segs) {
        drawSeg(canvas, seg, x + x0 + seg.x(), lineTop, base);
      }
      drawBorder(canvas, borderTop, left, top, left + width, top, 0, 1);
      drawBorder(canvas, borderBottom, left, top + boxHeight, left + width, top + boxHeight, 0,
          -1);
      drawBorder(canvas, borderLeft, left, top, left, top + boxHeight, -1, 0);
      drawBorder(canvas, borderRight, left + width, top, left + width, top + boxHeight, 1, 0);
      for (Model.Image image : floating) {
        drawFloating(canvas, image, x, lineTop);
      }
    }

    private void drawSeg(PageCanvas canvas, Seg seg, float sx, float lineTop, float base) {
      final Style st = seg.style();
      final float h = lineHeight;
      Integer highlight = st.highlight();
      if (highlight != null && seg.width() > 0) {
        canvas.add(p -> p.fillRect(sx, lineTop, seg.width(), h, highlight));
      }
      Model.Image image = seg.image();
      if (image != null) {
        canvas.add(p -> p.image(image.data(), sx, base - image.height(), image.width(),
            image.height()));
      }
      PDType0Font font = seg.font();
      if (font != null && !seg.text().isEmpty()) {
        canvas.add(p -> p.text(sx, base - seg.rise(), seg.text(), font, seg.size(),
            st.color(), seg.charSpacing(), seg.syntheticBold(), st.italic()));
      }
      if (seg.width() > 0) {
        float thickness = Math.max(0.4f, seg.size() / 18f);
        if (st.underline()) {
          float uy = base + seg.size() * 0.12f;
          canvas.add(p -> p.line(sx, uy, sx + seg.width(), uy, thickness, st.color(),
              "single"));
        }
        if (st.strike()) {
          float sy = base - seg.rise() - seg.size() * 0.28f;
          canvas.add(p -> p.line(sx, sy, sx + seg.width(), sy, thickness, st.color(),
              "single"));
        }
        String link = st.link();
        if (link != null) {
          canvas.add(p -> p.link(sx, lineTop, seg.width(), h, link));
        }
      }
    }

    /** Draws a border, moved outwards by its space ({@code dx}, {@code dy}: the direction). */
    private static void drawBorder(PageCanvas canvas, @Nullable Border b, float x1, float y1,
        float x2, float y2, int dx, int dy) {
      if (b == null || !b.visible()) {
        return;
      }
      // Top / bottom borders are inside the padding. Left / right ones go out into the indent.
      float ox = dx * (b.space() + b.width() / 2);
      float oy = dy * b.width() / 2;
      canvas.add(p -> p.line(x1 + ox, y1 + oy, x2 + ox, y2 + oy, b.width(), b.color(),
          b.style()));
    }

    private static void drawFloating(PageCanvas canvas, Model.Image image, float areaLeft,
        float lineTop) {
      Model.Anchor a = image.anchor();
      if (a == null) {
        return;
      }
      Model.Section s = canvas.section;
      float refX;
      float refW;
      switch (a.relH()) {
        case "page" -> {
          refX = 0;
          refW = s.pageWidth;
        }
        case "margin" -> {
          refX = s.marginLeft;
          refW = s.pageWidth - s.marginLeft - s.marginRight;
        }
        case "leftMargin", "insideMargin" -> {
          refX = 0;
          refW = s.marginLeft;
        }
        case "rightMargin", "outsideMargin" -> {
          refX = s.pageWidth - s.marginRight;
          refW = s.marginRight;
        }
        default -> {
          refX = areaLeft;
          refW = s.pageWidth - s.marginLeft - s.marginRight;
        }
      }
      float px = refX + a.offsetX();
      String alignH = a.alignH();
      if ("center".equals(alignH)) {
        px = refX + (refW - image.width()) / 2;
      } else if ("right".equals(alignH) || "outside".equals(alignH)) {
        px = refX + refW - image.width();
      } else if (alignH != null) {
        px = refX;
      }
      float refY;
      float refH;
      switch (a.relV()) {
        case "page" -> {
          refY = 0;
          refH = s.pageHeight;
        }
        case "margin" -> {
          refY = s.marginTop;
          refH = s.pageHeight - s.marginTop - s.marginBottom;
        }
        case "topMargin" -> {
          refY = 0;
          refH = s.marginTop;
        }
        case "bottomMargin" -> {
          refY = s.pageHeight - s.marginBottom;
          refH = s.marginBottom;
        }
        default -> {
          refY = lineTop;
          refH = 0;
        }
      }
      float py = refY + a.offsetY();
      String alignV = a.alignV();
      if ("center".equals(alignV)) {
        py = refY + (refH - image.height()) / 2;
      } else if ("bottom".equals(alignV) || "outside".equals(alignV)) {
        py = refY + refH - image.height();
      } else if (alignV != null) {
        py = refY;
      }
      final float fx = px;
      final float fy = py;
      canvas.add(p -> p.image(image.data(), fx, fy, image.width(), image.height()));
    }
  }
}
