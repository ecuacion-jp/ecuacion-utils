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
package jp.ecuacion.util.pdf.word.document.util;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * Builds a minimal .docx file from WordprocessingML fragments, so that tests can state exactly
 * the markup they verify.
 */
final class DocxBuilder {

  private static final String NS = "xmlns:w=\"http://schemas.openxmlformats.org/"
      + "wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/"
      + "officeDocument/2006/relationships\"";
  private static final String REL =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  private static final String XML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";

  private final StringBuilder body = new StringBuilder();
  private String sectPr = "<w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" "
      + "w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/>";
  private final Map<String, String> parts = new LinkedHashMap<>();
  private final Map<String, byte[]> media = new LinkedHashMap<>();
  private final StringBuilder rels = new StringBuilder();

  /** Returns a run with the text. */
  static String r(String text) {
    return r(text, "");
  }

  /** Returns a run with the text and the run properties. */
  static String r(String text, String rpr) {
    return "<w:r><w:rPr>" + rpr + "</w:rPr><w:t xml:space=\"preserve\">" + text
        + "</w:t></w:r>";
  }

  /** Returns a paragraph. */
  static String p(String inner) {
    return p("", inner);
  }

  /** Returns a paragraph with the paragraph properties. */
  static String p(String ppr, String inner) {
    return "<w:p><w:pPr>" + ppr + "</w:pPr>" + inner + "</w:p>";
  }

  /** Returns a complex field with the cached result. */
  static String field(String instruction, String cached) {
    return "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> " + instruction
        + " </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + r(cached) + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
  }

  /** Returns a run which breaks the page. */
  static String pageBreak() {
    return "<w:r><w:br w:type=\"page\"/></w:r>";
  }

  DocxBuilder add(String blockXml) {
    body.append(blockXml);
    return this;
  }

  /** Replaces the section properties of the (last) section. */
  DocxBuilder sectPr(String xml) {
    this.sectPr = xml;
    return this;
  }

  /** Adds content to the section properties of the (last) section. */
  DocxBuilder addToSectPr(String xml) {
    this.sectPr += xml;
    return this;
  }

  DocxBuilder styles(String stylesInner) {
    return part("styles.xml", "styles", "<w:styles " + NS + ">" + stylesInner + "</w:styles>",
        "rIdStyles");
  }

  DocxBuilder numbering(String numberingInner) {
    return part("numbering.xml", "numbering",
        "<w:numbering " + NS + ">" + numberingInner + "</w:numbering>", "rIdNumbering");
  }

  DocxBuilder settings(String settingsInner) {
    return part("settings.xml", "settings",
        "<w:settings " + NS + ">" + settingsInner + "</w:settings>", "rIdSettings");
  }

  /** Adds a default header with the content. */
  DocxBuilder header(String blocks) {
    addToSectPr("<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>");
    return part("header1.xml", "header", "<w:hdr " + NS + ">" + blocks + "</w:hdr>",
        "rIdHeader");
  }

  /** Adds a default footer with the content. */
  DocxBuilder footer(String blocks) {
    addToSectPr("<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>");
    return part("footer1.xml", "footer", "<w:ftr " + NS + ">" + blocks + "</w:ftr>",
        "rIdFooter");
  }

  /**
   * Adds a header or footer part which section properties refer to by the relationship ID
   * (e.g. {@code <w:headerReference w:type="first" r:id="..."/>}).
   */
  DocxBuilder story(String relId, boolean header, String blocks) {
    String tag = header ? "hdr" : "ftr";
    return part(relId + ".xml", header ? "header" : "footer",
        "<w:" + tag + " " + NS + ">" + blocks + "</w:" + tag + ">", relId);
  }

  /** Adds a PNG picture which drawings refer to by the relationship ID. */
  DocxBuilder png(String relId, byte[] bytes) {
    media.put("word/media/" + relId + ".png", bytes.clone());
    rels.append("<Relationship Id=\"").append(relId).append("\" Type=\"").append(REL)
        .append("/image\" Target=\"media/").append(relId).append(".png\"/>");
    return this;
  }

  /** Returns a run with an inline picture of the size in points. */
  static String inlinePicture(String relId, int width, int height) {
    return "<w:r><w:drawing><wp:inline " + DRAWING_NS + ">" + extentAndGraphic(relId, width,
        height) + "</wp:inline></w:drawing></w:r>";
  }

  /**
   * Returns a run with a floating picture, at the offsets in points from what
   * {@code relativeH} / {@code relativeV} say (page, margin, column, paragraph etc.).
   */
  static String floatingPicture(String relId, int width, int height, String relativeH,
      int offsetX, String relativeV, int offsetY) {
    return "<w:r><w:drawing><wp:anchor " + DRAWING_NS + " behindDoc=\"0\" simplePos=\"0\" "
        + "relativeHeight=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
        + "<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"" + relativeH
        + "\"><wp:posOffset>" + offsetX * 12700 + "</wp:posOffset></wp:positionH>"
        + "<wp:positionV relativeFrom=\"" + relativeV + "\"><wp:posOffset>" + offsetY * 12700
        + "</wp:posOffset></wp:positionV>" + extentAndGraphic(relId, width, height).replace(
            "<wp:docPr", "<wp:wrapNone/><wp:docPr")
        + "</wp:anchor></w:drawing></w:r>";
  }

  private static final String DRAWING_NS = "xmlns:wp=\"http://schemas.openxmlformats.org/"
      + "drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/"
      + "drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/"
      + "picture\"";

  private static String extentAndGraphic(String relId, int width, int height) {
    String size = "cx=\"" + width * 12700 + "\" cy=\"" + height * 12700 + "\"";
    return "<wp:extent " + size + "/><wp:docPr id=\"1\" name=\"p\"/><a:graphic>"
        + "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
        + "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr>"
        + "<pic:blipFill><a:blip r:embed=\"" + relId + "\"/><a:stretch><a:fillRect/>"
        + "</a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext " + size
        + "/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>"
        + "</a:graphicData></a:graphic>";
  }

  private DocxBuilder part(String name, String type, String xml, String relId) {
    parts.put("word/" + name, XML + xml);
    rels.append("<Relationship Id=\"").append(relId).append("\" Type=\"").append(REL)
        .append('/').append(type).append("\" Target=\"").append(name).append("\"/>");
    return this;
  }

  /** Writes the .docx file. */
  Path write(Path path) throws IOException {
    String document = XML + "<w:document " + NS + "><w:body>" + body + "<w:sectPr>" + sectPr
        + "</w:sectPr></w:body></w:document>";
    String contentTypes = XML + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/"
        + "content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd."
        + "openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" "
        + "ContentType=\"application/xml\"/><Default Extension=\"png\" "
        + "ContentType=\"image/png\"/><Override PartName=\"/word/document.xml\" "
        + "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml."
        + "document.main+xml\"/></Types>";
    String packageRels = XML + "<Relationships xmlns=\"http://schemas.openxmlformats.org/"
        + "package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + REL
        + "/officeDocument\" Target=\"word/document.xml\"/></Relationships>";
    String documentRels = XML + "<Relationships xmlns=\"http://schemas.openxmlformats.org/"
        + "package/2006/relationships\">" + rels + "</Relationships>";
    try (OutputStream os = Files.newOutputStream(path);
        ZipOutputStream zip = new ZipOutputStream(os)) {
      put(zip, "[Content_Types].xml", contentTypes);
      put(zip, "_rels/.rels", packageRels);
      put(zip, "word/_rels/document.xml.rels", documentRels);
      put(zip, "word/document.xml", document);
      for (Map.Entry<String, String> e : parts.entrySet()) {
        put(zip, e.getKey(), e.getValue());
      }
      for (Map.Entry<String, byte[]> e : media.entrySet()) {
        zip.putNextEntry(new ZipEntry(e.getKey()));
        zip.write(e.getValue());
        zip.closeEntry();
      }
    }
    return path;
  }

  private static void put(ZipOutputStream zip, String name, String content) throws IOException {
    zip.putNextEntry(new ZipEntry(name));
    zip.write(content.getBytes(StandardCharsets.UTF_8));
    zip.closeEntry();
  }
}
