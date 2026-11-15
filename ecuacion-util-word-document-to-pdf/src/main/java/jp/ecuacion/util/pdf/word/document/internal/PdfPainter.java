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
import java.util.HashSet;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Set;
import jp.ecuacion.lib.core.logging.DetailLogger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.RenderingMode;
import org.apache.pdfbox.pdmodel.interactive.action.PDActionGoTo;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationLink;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDBorderStyleDictionary;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.destination.PDPageXYZDestination;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDDocumentOutline;
import org.apache.pdfbox.pdmodel.interactive.documentnavigation.outline.PDOutlineItem;
import org.apache.pdfbox.util.Matrix;
import org.jspecify.annotations.Nullable;

/** Paints the result of the layout to a PDF document. Coordinates are from the top-left. */
final class PdfPainter {

  private static final DetailLogger detailLog = new DetailLogger(PdfPainter.class);

  private final PDDocument document;
  private final Layout.PassResult result;
  private final List<PDPage> pages = new ArrayList<>();
  private final IdentityHashMap<Model.ImageData, PDImageXObject> images =
      new IdentityHashMap<>();
  private final Set<Model.ImageData> brokenImages = new HashSet<>();
  private @Nullable PDPageContentStream stream;
  private @Nullable PDPage page;
  private float pageHeight;

  PdfPainter(PDDocument document, Layout.PassResult result) {
    this.document = document;
    this.result = result;
  }

  void paint() throws IOException {
    for (Layout.PageCanvas canvas : result.pages) {
      PDPage p =
          new PDPage(new PDRectangle(canvas.section.pageWidth, canvas.section.pageHeight));
      document.addPage(p);
      pages.add(p);
    }
    for (int i = 0; i < pages.size(); i++) {
      Layout.PageCanvas canvas = result.pages.get(i);
      page = pages.get(i);
      pageHeight = canvas.section.pageHeight;
      try (PDPageContentStream cs = new PDPageContentStream(document, pages.get(i))) {
        stream = cs;
        for (Layout.Op op : canvas.ops) {
          op.paint(this);
        }
      }
    }
    stream = null;
    paintOutline();
  }

  private PDPageContentStream cs() {
    PDPageContentStream cs = stream;
    if (cs == null) {
      throw new IllegalStateException("No page is being painted.");
    }
    return cs;
  }

  private static float[] rgb(int color) {
    int c = color < 0 ? 0 : color;
    return new float[] {((c >> 16) & 0xFF) / 255f, ((c >> 8) & 0xFF) / 255f, (c & 0xFF) / 255f};
  }

  void fillRect(float x, float y, float w, float h, int color) throws IOException {
    if (color < 0 || w <= 0 || h <= 0) {
      return;
    }
    PDPageContentStream cs = cs();
    float[] c = rgb(color);
    cs.setNonStrokingColor(c[0], c[1], c[2]);
    cs.addRect(x, pageHeight - y - h, w, h);
    cs.fill();
  }

  void line(float x1, float y1, float x2, float y2, float width, int color, String style)
      throws IOException {
    PDPageContentStream cs = cs();
    float[] c = rgb(color);
    cs.setStrokingColor(c[0], c[1], c[2]);
    if ("double".equals(style)) {
      // Two lines of the width, one line width apart.
      boolean horizontal = Math.abs(y1 - y2) < Math.abs(x1 - x2);
      float d = width;
      cs.setLineWidth(d);
      for (int sign = -1; sign <= 1; sign += 2) {
        float ox = horizontal ? 0 : sign * d;
        float oy = horizontal ? sign * d : 0;
        cs.moveTo(x1 + ox, pageHeight - (y1 + oy));
        cs.lineTo(x2 + ox, pageHeight - (y2 + oy));
        cs.stroke();
      }
      return;
    }
    cs.setLineWidth(width);
    boolean dashed = style.startsWith("dash");
    boolean dotted = style.startsWith("dot");
    if (dashed) {
      cs.setLineDashPattern(new float[] {width * 3, width * 3}, 0);
    } else if (dotted) {
      cs.setLineDashPattern(new float[] {width, width}, 0);
    }
    cs.moveTo(x1, pageHeight - y1);
    cs.lineTo(x2, pageHeight - y2);
    cs.stroke();
    if (dashed || dotted) {
      cs.setLineDashPattern(new float[0], 0);
    }
  }

  void text(float x, float baseline, String text, PDType0Font font, float size, int color,
      float charSpacing, boolean syntheticBold, boolean italic) throws IOException {
    PDPageContentStream cs = cs();
    float[] c = rgb(color);
    cs.saveGraphicsState();
    cs.setNonStrokingColor(c[0], c[1], c[2]);
    if (syntheticBold) {
      cs.setStrokingColor(c[0], c[1], c[2]);
      cs.setLineWidth(size / 36f);
    }
    cs.beginText();
    cs.setFont(font, size);
    if (charSpacing != 0f) {
      cs.setCharacterSpacing(charSpacing);
    }
    if (syntheticBold) {
      cs.setRenderingMode(RenderingMode.FILL_STROKE);
    }
    float py = pageHeight - baseline;
    cs.setTextMatrix(italic ? new Matrix(1, 0, 0.2f, 1, x, py)
        : Matrix.getTranslateInstance(x, py));
    cs.showText(text);
    cs.endText();
    cs.restoreGraphicsState();
  }

  void image(Model.ImageData data, float x, float y, float w, float h) throws IOException {
    if (brokenImages.contains(data)) {
      return;
    }
    PDImageXObject image = images.get(data);
    if (image == null) {
      try {
        image = PDImageXObject.createFromByteArray(document, data.bytes(), null);
      } catch (IOException | IllegalArgumentException e) {
        brokenImages.add(data);
        detailLog.warn("A picture in the Word file is in a format which is not supported "
            + "(e.g. EMF / WMF), so it is not output to the PDF file: " + e.getMessage());
        return;
      }
      images.put(data, image);
    }
    cs().drawImage(image, x, pageHeight - y - h, w, h);
  }

  void link(float x, float y, float w, float h, String anchor) throws IOException {
    float[] target = result.bookmarkPositions.get(anchor);
    PDPage current = page;
    if (target == null || current == null || (int) target[0] >= pages.size()) {
      return;
    }
    PDAnnotationLink link = new PDAnnotationLink();
    PDBorderStyleDictionary border = new PDBorderStyleDictionary();
    border.setWidth(0);
    link.setBorderStyle(border);
    link.setRectangle(new PDRectangle(x, pageHeight - y - h, w, h));
    PDActionGoTo action = new PDActionGoTo();
    action.setDestination(destination((int) target[0], target[1]));
    link.setAction(action);
    current.getAnnotations().add(link);
  }

  private PDPageXYZDestination destination(int pageIndex, float y) {
    PDPage target = pages.get(pageIndex);
    PDPageXYZDestination dest = new PDPageXYZDestination();
    dest.setPage(target);
    dest.setLeft(0);
    dest.setTop(Math.round(target.getMediaBox().getHeight() - y));
    return dest;
  }

  /** Adds the headings as the outline (bookmarks) of the PDF document. */
  private void paintOutline() {
    if (result.outlines.isEmpty()) {
      return;
    }
    PDDocumentOutline outline = new PDDocumentOutline();
    document.getDocumentCatalog().setDocumentOutline(outline);
    List<PDOutlineItem> parents = new ArrayList<>();
    List<Integer> levels = new ArrayList<>();
    for (Layout.Outline o : result.outlines) {
      if (o.page() >= pages.size()) {
        continue;
      }
      PDOutlineItem item = new PDOutlineItem();
      item.setTitle(o.title());
      item.setDestination(destination(o.page(), o.y()));
      while (!levels.isEmpty() && levels.get(levels.size() - 1) >= o.level()) {
        levels.removeLast();
        parents.removeLast();
      }
      if (parents.isEmpty()) {
        outline.addLast(item);
      } else {
        parents.get(parents.size() - 1).addLast(item);
      }
      parents.add(item);
      levels.add(o.level());
    }
  }
}
