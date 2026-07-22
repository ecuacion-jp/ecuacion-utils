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
package jp.ecuacion.util.pdf.excel.report.internal;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.JPEGFactory;
import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGeomGuide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;

/** Renders Excel shapes (images and auto-shapes) onto PDF pages. */
class ShapeRenderer {

  private static final float PX_TO_PT = 72f / 96f;
  private static final float CELL_PADDING = 2f;

  /**
   * Upper bound of embedded-image pixel area (width × height) rendered into the PDF.
   * 50M pixels ≈ an A4 page at 600 DPI; anything larger is treated as a decompression bomb.
   */
  private static final long MAX_IMAGE_PIXELS = 50_000_000L;

  private final PDDocument document;
  private final FontManager fontManager;

  ShapeRenderer(PDDocument document, FontManager fontManager) {
    this.document = document;
    this.fontManager = fontManager;
  }

  void renderShapes(Sheet sheet, PDPageContentStream cs, PDRectangle pageSize,
      float leftMargin, float topMargin, int firstPageRow, int lastPageRow, int firstPageCol,
      int printFirstRow, int printFirstCol, float[] rowHeights, float[] colWidths,
      float scaleFactor) throws IOException {

    if (!(sheet instanceof XSSFSheet xssfSheet)) {
      return;
    }
    XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
    if (drawing == null) {
      return;
    }

    float pageTopExcelY = 0f;
    for (int r = printFirstRow; r < firstPageRow; r++) {
      pageTopExcelY += rowHeights[r - printFirstRow];
    }
    float pageBottomExcelY = pageTopExcelY;
    for (int r = firstPageRow; r <= lastPageRow; r++) {
      pageBottomExcelY += rowHeights[r - printFirstRow];
    }

    float pageLeftExcelX = 0f;
    for (int c = printFirstCol; c < firstPageCol; c++) {
      pageLeftExcelX += colWidths[c - printFirstCol];
    }

    int printLastCol = printFirstCol + colWidths.length - 1;
    int printLastRow = printFirstRow + rowHeights.length - 1;

    for (XSSFShape shape : drawing.getShapes()) {
      if (!(shape.getAnchor() instanceof XSSFClientAnchor anchor)) {
        continue;
      }

      // Skip shapes whose anchor start column or row is outside the print area.
      if (anchor.getCol1() < printFirstCol || anchor.getCol1() > printLastCol) {
        continue;
      }
      if (anchor.getRow1() < printFirstRow || anchor.getRow1() > printLastRow) {
        continue;
      }

      float shapeTopY = shapeExcelY(anchor.getRow1(), anchor.getDy1(), rowHeights, printFirstRow);
      float shapeLeftX = shapeExcelX(anchor.getCol1(), anchor.getDx1(), colWidths, printFirstCol);
      float shapeBottomY;
      float shapeRightX;

      if (anchor.getRow2() > 0 || anchor.getCol2() > 0) {
        shapeBottomY = shapeExcelY(anchor.getRow2(), anchor.getDy2(), rowHeights, printFirstRow);
        shapeRightX = shapeExcelX(anchor.getCol2(), anchor.getDx2(), colWidths, printFirstCol);
      } else if (shape instanceof XSSFPicture picture) {
        Dimension dim = picture.getImageDimension();
        if (dim == null || dim.width <= 0 || dim.height <= 0) {
          continue;
        }
        shapeBottomY = shapeTopY + dim.height * PX_TO_PT * scaleFactor;
        shapeRightX = shapeLeftX + dim.width * PX_TO_PT * scaleFactor;
      } else if (shape instanceof XSSFSimpleShape simpleShape) {
        long extCx = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCx();
        long extCy = simpleShape.getCTShape().getSpPr().getXfrm().getExt().getCy();
        shapeBottomY = shapeTopY + extCy / 12700f * scaleFactor;
        shapeRightX = shapeLeftX + extCx / 12700f * scaleFactor;
      } else {
        continue;
      }

      if (shapeTopY < pageTopExcelY || shapeTopY >= pageBottomExcelY) {
        continue;
      }

      float relTop = shapeTopY - pageTopExcelY;
      float relBottom = shapeBottomY - pageTopExcelY;
      float relLeft = shapeLeftX - pageLeftExcelX;

      float pdfShapeBottom = pageSize.getHeight() - topMargin - relBottom;
      float pdfShapeHeight = relBottom - relTop;
      float pdfShapeLeft = leftMargin + relLeft;
      float pdfShapeWidth = shapeRightX - shapeLeftX;

      if (shape instanceof XSSFPicture picture) {
        renderPicture(cs, picture, pdfShapeLeft, pdfShapeBottom, pdfShapeWidth, pdfShapeHeight);
      } else if (shape instanceof XSSFSimpleShape simpleShape) {
        renderShape(cs, simpleShape, pdfShapeLeft, pdfShapeBottom, pdfShapeWidth, pdfShapeHeight,
            scaleFactor);
      }
    }
  }

  /** Converts a shape anchor's row + EMU offset to an Excel-coordinate Y value. */
  private float shapeExcelY(int row, int dyEmu, float[] rowHeights, int printFirstRow) {
    float y = 0f;
    int limit = Math.min(row, printFirstRow + rowHeights.length);
    for (int r = printFirstRow; r < limit; r++) {
      y += rowHeights[r - printFirstRow];
    }
    return y + dyEmu / 12700f;
  }

  /** Converts a shape anchor's column + EMU offset to an Excel-coordinate X value. */
  private float shapeExcelX(int col, int dxEmu, float[] colWidths, int printFirstCol) {
    float x = 0f;
    int limit = Math.min(col, printFirstCol + colWidths.length);
    for (int c = printFirstCol; c < limit; c++) {
      x += colWidths[c - printFirstCol];
    }
    return x + dxEmu / 12700f;
  }

  private void renderPicture(PDPageContentStream cs, XSSFPicture picture, float x, float y,
      float width, float height) throws IOException {
    XSSFPictureData picData = picture.getPictureData();
    byte[] imageBytes = picData.getData();
    if (exceedsPixelLimit(imageBytes)) {
      return;
    }
    String mime = picData.getMimeType();
    PDImageXObject pdImage;
    if ("image/jpeg".equalsIgnoreCase(mime) || "image/jpg".equalsIgnoreCase(mime)) {
      pdImage = JPEGFactory.createFromByteArray(document, imageBytes);
    } else {
      BufferedImage bi = ImageIO.read(new ByteArrayInputStream(imageBytes));
      if (bi == null) {
        return;
      }
      pdImage = LosslessFactory.createFromImage(document, bi);
    }
    cs.drawImage(pdImage, x, y, width, height);
  }

  /**
   * Returns {@code true} when the image's declared dimensions exceed {@link #MAX_IMAGE_PIXELS}.
   *
   * <p>A decompression bomb (a small file declaring a huge pixel area) would otherwise make
   * {@code ImageIO.read} allocate gigabytes of heap. The dimensions are read from the image
   * header only, without decoding pixel data.</p>
   */
  private static boolean exceedsPixelLimit(byte[] imageBytes) {
    try (var iis = ImageIO.createImageInputStream(new ByteArrayInputStream(imageBytes))) {
      var readers = ImageIO.getImageReaders(iis);
      if (!readers.hasNext()) {
        return false;
      }
      var reader = readers.next();
      try {
        reader.setInput(iis, true, true);
        long pixels = (long) reader.getWidth(0) * reader.getHeight(0);
        return pixels > MAX_IMAGE_PIXELS;
      } finally {
        reader.dispose();
      }
    } catch (IOException e) {
      // Let the subsequent decode report the failure in its usual way.
      return false;
    }
  }

  private void renderShape(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

    Color fillColor = getShapeFillColor(shape);
    Color lineColor = getShapeLineColor(shape);
    float lineWidth = getShapeLineWidth(shape);
    boolean hasLine = lineColor != null && lineWidth > 0f;

    if (fillColor != null && hasLine) {
      cs.setNonStrokingColor(fillColor);
      cs.setStrokingColor(lineColor);
      cs.setLineWidth(lineWidth);
      appendShapePath(cs, shape, x, y, width, height);
      cs.fillAndStroke();
    } else if (fillColor != null) {
      cs.setNonStrokingColor(fillColor);
      appendShapePath(cs, shape, x, y, width, height);
      cs.fill();
    } else if (hasLine) {
      cs.setStrokingColor(lineColor);
      cs.setLineWidth(lineWidth);
      appendShapePath(cs, shape, x, y, width, height);
      cs.stroke();
    }

    renderShapeText(cs, shape, x, y, width, height, scaleFactor);
  }

  private void appendShapePath(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height) throws IOException {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    STShapeType.Enum shapeType =
        (spPr != null && spPr.isSetPrstGeom()) ? spPr.getPrstGeom().getPrst() : null;

    if (STShapeType.PARALLELOGRAM == shapeType) {
      float offset = (float) (readAdj(spPr, 0.25) * height);
      cs.moveTo(x + offset, y + height);
      cs.lineTo(x + width, y + height);
      cs.lineTo(x + width - offset, y);
      cs.lineTo(x, y);
      cs.closePath();
    } else if (STShapeType.DIAMOND == shapeType) {
      float cx = x + width / 2f;
      float cy = y + height / 2f;
      cs.moveTo(cx, y + height);
      cs.lineTo(x + width, cy);
      cs.lineTo(cx, y);
      cs.lineTo(x, cy);
      cs.closePath();
    } else if (STShapeType.ROUND_RECT == shapeType) {
      float r = (float) (readAdj(spPr, 0.16667) * Math.min(width, height));
      appendRoundRectPath(cs, x, y, width, height, r);
    } else if (STShapeType.ELLIPSE == shapeType) {
      float kappa = 0.5523f;
      float rx = width / 2f;
      float ry = height / 2f;
      float cx = x + rx;
      float cy = y + ry;
      cs.moveTo(cx, cy + ry);
      cs.curveTo(cx + rx * kappa, cy + ry, cx + rx, cy + ry * kappa, cx + rx, cy);
      cs.curveTo(cx + rx, cy - ry * kappa, cx + rx * kappa, cy - ry, cx, cy - ry);
      cs.curveTo(cx - rx * kappa, cy - ry, cx - rx, cy - ry * kappa, cx - rx, cy);
      cs.curveTo(cx - rx, cy + ry * kappa, cx - rx * kappa, cy + ry, cx, cy + ry);
      cs.closePath();
    } else {
      cs.addRect(x, y, width, height);
    }
  }

  private void appendRoundRectPath(PDPageContentStream cs, float x, float y, float width,
      float height, float r) throws IOException {
    float k = 0.5523f * r;
    cs.moveTo(x + r, y);
    cs.lineTo(x + width - r, y);
    cs.curveTo(x + width - r + k, y, x + width, y + k, x + width, y + r);
    cs.lineTo(x + width, y + height - r);
    cs.curveTo(x + width, y + height - r + k, x + width - r + k, y + height, x + width - r,
        y + height);
    cs.lineTo(x + r, y + height);
    cs.curveTo(x + r - k, y + height, x, y + height - r + k, x, y + height - r);
    cs.lineTo(x, y + r);
    cs.curveTo(x, y + r - k, x + r - k, y, x + r, y);
    cs.closePath();
  }

  private double readAdj(CTShapeProperties spPr, double defaultVal) {
    if (!spPr.isSetPrstGeom()) {
      return defaultVal;
    }
    CTPresetGeometry2D prstGeom = spPr.getPrstGeom();
    if (!prstGeom.isSetAvLst()) {
      return defaultVal;
    }
    for (CTGeomGuide gd : prstGeom.getAvLst().getGdList()) {
      if ("adj".equals(gd.getName()) && gd.getFmla() != null) {
        String fmla = gd.getFmla().trim();
        if (fmla.startsWith("val ")) {
          try {
            return Long.parseLong(fmla.substring(4)) / 100000.0;
          } catch (NumberFormatException e) {
            // fall through to default
          }
        }
      }
    }
    return defaultVal;
  }

  private @Nullable Color getShapeFillColor(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || spPr.isSetNoFill() || !spPr.isSetSolidFill()) {
      return null;
    }
    CTSolidColorFillProperties fill = spPr.getSolidFill();
    if (!fill.isSetSrgbClr()) {
      return null;
    }
    byte[] rgb = fill.getSrgbClr().getVal();
    return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
  }

  private @Nullable Color getShapeLineColor(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || !spPr.isSetLn()) {
      return null;
    }
    CTLineProperties ln = spPr.getLn();
    if (ln.isSetNoFill() || !ln.isSetSolidFill()) {
      return null;
    }
    CTSolidColorFillProperties fill = ln.getSolidFill();
    if (!fill.isSetSrgbClr()) {
      return null;
    }
    byte[] rgb = fill.getSrgbClr().getVal();
    return new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
  }

  private float getShapeLineWidth(XSSFSimpleShape shape) {
    CTShapeProperties spPr = shape.getCTShape().getSpPr();
    if (spPr == null || !spPr.isSetLn()) {
      return 0f;
    }
    CTLineProperties ln = spPr.getLn();
    if (!ln.isSetW()) {
      return 0f;
    }
    return ln.getW() / 12700f;
  }

  private void renderShapeText(PDPageContentStream cs, XSSFSimpleShape shape, float x, float y,
      float width, float height, float scaleFactor) throws IOException {

    List<XSSFTextParagraph> paragraphs = shape.getTextParagraphs();
    if (paragraphs.isEmpty()) {
      return;
    }

    var txBody = shape.getCTShape().isSetTxBody() ? shape.getCTShape().getTxBody() : null;

    float leftInset = CELL_PADDING * scaleFactor;
    float rightInset = CELL_PADDING * scaleFactor;
    float topInset = CELL_PADDING * scaleFactor;
    var bodyPr = (txBody != null) ? txBody.getBodyPr() : null;
    if (bodyPr != null) {
      long leftIns = 91440L;
      long rightIns = 91440L;
      long topIns = 45720L;
      if (bodyPr.isSetLIns()) {
        leftIns = ((Number) bodyPr.getLIns()).longValue();
      }
      if (bodyPr.isSetRIns()) {
        rightIns = ((Number) bodyPr.getRIns()).longValue();
      }
      if (bodyPr.isSetTIns()) {
        topIns = ((Number) bodyPr.getTIns()).longValue();
      }
      leftInset = leftIns / 12700f * scaleFactor;
      rightInset = rightIns / 12700f * scaleFactor;
      topInset = topIns / 12700f * scaleFactor;
    }

    float curY = y + height - topInset;

    for (int paraIdx = 0; paraIdx < paragraphs.size(); paraIdx++) {
      XSSFTextParagraph para = paragraphs.get(paraIdx);
      List<XSSFTextRun> runs = para.getTextRuns();
      if (runs.isEmpty()) {
        continue;
      }

      StringBuilder sb = new StringBuilder();
      for (XSSFTextRun run : runs) {
        sb.append(run.getText());
      }
      String lineText = sb.toString();
      if (lineText.isBlank()) {
        continue;
      }

      XSSFTextRun firstRun = runs.get(0);
      double fontSizePt = firstRun.getFontSize();
      if (fontSizePt <= 0) {
        fontSizePt = 11.0;
      }
      float fontSize = (float) fontSizePt * scaleFactor;
      boolean bold = firstRun.isBold();
      PDType0Font font = fontManager.getFont(bold);

      Color textColor = firstRun.getFontColor();
      if (textColor == null) {
        textColor = Color.BLACK;
      }

      float textWidth;
      try {
        textWidth = font.getStringWidth(lineText) / 1000f * fontSize;
      } catch (Exception e) {
        textWidth = width;
      }

      org.apache.poi.xssf.usermodel.TextAlign align = para.getTextAlign();
      var ctPara =
          (txBody != null && paraIdx < txBody.sizeOfPArray()) ? txBody.getPArray(paraIdx) : null;
      boolean paraRtl = ctPara != null && ctPara.isSetPPr() && ctPara.getPPr().isSetRtl()
          && ctPara.getPPr().getRtl();
      if (paraRtl && align == org.apache.poi.xssf.usermodel.TextAlign.LEFT) {
        align = org.apache.poi.xssf.usermodel.TextAlign.CENTER;
      }

      float textX;
      if (align == org.apache.poi.xssf.usermodel.TextAlign.CENTER) {
        float availWidth = width - leftInset - rightInset;
        textX = x + leftInset + Math.max(0f, (availWidth - textWidth) / 2f);
      } else if (align == org.apache.poi.xssf.usermodel.TextAlign.RIGHT) {
        textX = x + width - rightInset - textWidth;
      } else {
        textX = x + leftInset;
      }

      final float ascent = font.getFontDescriptor().getAscent() / 1000f * fontSize;
      final float descent = font.getFontDescriptor().getDescent() / 1000f * fontSize;
      curY -= ascent;

      cs.beginText();
      cs.setFont(font, fontSize);
      cs.setNonStrokingColor(textColor);
      cs.newLineAtOffset(textX, curY);
      try {
        cs.showText(lineText);
      } catch (Exception e) {
        // Skip text that cannot be rendered with the current font
      }
      cs.endText();

      curY += descent;
      curY -= CELL_PADDING * scaleFactor;
    }
  }
}
