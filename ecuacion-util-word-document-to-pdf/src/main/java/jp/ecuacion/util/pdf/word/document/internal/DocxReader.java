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
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import jp.ecuacion.lib.core.logging.DetailLogger;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/** Reads a .docx package into a {@link Model.Document}. */
final class DocxReader {

  private static final DetailLogger detailLog = new DetailLogger(DocxReader.class);

  private static final Set<String> DYNAMIC_FIELDS =
      Set.of("PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF");
  private static final float EMU_PER_POINT = 12700f;

  /** The table style formatting which applies to the paragraphs in a cell. */
  private record TableContext(ParaProps ppr, RunProps rpr) {
  }

  private static final class FieldState {
    final StringBuilder instruction = new StringBuilder();
    final StringBuilder fallback = new StringBuilder();
    boolean parsed;
    boolean inResult;
    boolean dynamic;
    String kind = "";
    String arg = "";
    String format = "";
    @Nullable RunProps rpr;
    @Nullable String link;
  }

  private final DocxPackage pkg;
  private final Theme theme;
  private final StyleSheet styles;
  private final Numbering numbering;
  private final Model.Settings settings = new Model.Settings();
  private final Map<PackagePart, Map<String, Model.ImageData>> images = new HashMap<>();

  private PackagePart part;
  private Deque<FieldState> fields = new ArrayDeque<>();
  private @Nullable String hyperlink;
  private final List<String> pendingBookmarks = new ArrayList<>();
  private boolean warnedUnsupportedDrawing;

  private DocxReader(DocxPackage pkg) throws IOException {
    this.pkg = pkg;
    this.part = pkg.mainPart();
    String eastAsiaLang = readSettings(root(pkg.relatedByType(part, "/settings")));
    this.theme = Theme.parse(root(pkg.relatedByType(part, "/theme")), eastAsiaLang);
    this.styles = new StyleSheet(root(pkg.relatedByType(part, "/styles")), theme);
    this.numbering = new Numbering(root(pkg.relatedByType(part, "/numbering")), theme);
  }

  static Model.Document read(DocxPackage pkg) throws IOException {
    return new DocxReader(pkg).readDocument();
  }

  private @Nullable Element root(@Nullable PackagePart p) throws IOException {
    return p == null ? null : pkg.root(p);
  }

  private @Nullable String readSettings(@Nullable Element root) {
    Float tab = Xml.twips(Xml.child(root, "defaultTabStop"), "val");
    if (tab != null && tab > 0) {
      settings.defaultTab = tab;
    }
    settings.evenAndOddHeaders = Boolean.TRUE.equals(Xml.onOff(root, "evenAndOddHeaders"));
    String csc = Xml.val(root, "characterSpacingControl");
    settings.compressPunctuation = csc != null && csc.startsWith("compressPunctuation");
    for (Element cs : Xml.children(Xml.child(root, "compat"), "compatSetting")) {
      if ("compatibilityMode".equals(Xml.attr(cs, "name"))) {
        Integer mode = Xml.intAttr(cs, "val");
        if (mode != null) {
          settings.compatibilityMode = mode;
        }
      }
    }
    // Since Word 2013 (compatibility mode 15) the lines in table cells never snap to the grid.
    settings.snapToGridInCell = settings.compatibilityMode < 15
        && !Boolean.TRUE.equals(Xml.onOff(Xml.child(root, "compat"), "doNotSnapToGridInCell"));
    return Xml.attr(Xml.child(root, "themeFontLang"), "eastAsia");
  }

  private Model.Document readDocument() throws IOException {
    Model.Document doc = new Model.Document();
    doc.settings = settings;
    doc.defaultRun = styles.runBase(null, null);
    Element body = Xml.child(pkg.root(pkg.mainPart()), "body");
    List<Model.Block> blocks = readBlocks(body, null);

    Model.Section previous = null;
    List<Model.Block> current = new ArrayList<>();
    for (Model.Block b : blocks) {
      current.add(b);
      if (b instanceof Model.Para p && p.sectPr != null) {
        previous = readSection(p.sectPr, previous);
        previous.blocks.addAll(current);
        doc.sections.add(previous);
        current = new ArrayList<>();
      }
    }
    Model.Section last = readSection(Xml.child(body, "sectPr"), previous);
    last.blocks.addAll(current);
    doc.sections.add(last);
    return doc;
  }

  private Model.Section readSection(@Nullable Element sectPr, Model.@Nullable Section previous)
      throws IOException {
    Model.Section s = new Model.Section();
    if (previous != null) {
      s.headers.putAll(previous.headers);
      s.footers.putAll(previous.footers);
    }
    if (sectPr == null) {
      return s;
    }
    Element pgSz = Xml.child(sectPr, "pgSz");
    s.pageWidth = or(Xml.twips(pgSz, "w"), s.pageWidth);
    s.pageHeight = or(Xml.twips(pgSz, "h"), s.pageHeight);
    Element pgMar = Xml.child(sectPr, "pgMar");
    s.marginTop = Math.abs(or(Xml.twips(pgMar, "top"), s.marginTop));
    s.marginBottom = Math.abs(or(Xml.twips(pgMar, "bottom"), s.marginBottom));
    float gutter = or(Xml.twips(pgMar, "gutter"), 0f);
    s.marginLeft = or(Xml.twips(pgMar, "left"), s.marginLeft) + gutter;
    s.marginRight = or(Xml.twips(pgMar, "right"), s.marginRight);
    s.headerDistance = or(Xml.twips(pgMar, "header"), s.headerDistance);
    s.footerDistance = or(Xml.twips(pgMar, "footer"), s.footerDistance);
    Element grid = Xml.child(sectPr, "docGrid");
    String gridType = Xml.attr(grid, "type");
    if (gridType != null && !"default".equals(gridType)) {
      if (!"snapToChars".equals(gridType)) {
        s.linePitch = or(Xml.twips(grid, "linePitch"), 0f);
      }
      Integer charSpace = Xml.intAttr(grid, "charSpace");
      if (charSpace != null && !"lines".equals(gridType)) {
        s.charDelta = charSpace / 4096f;
      }
    }
    s.titlePage = Boolean.TRUE.equals(Xml.onOff(sectPr, "titlePg"));
    String type = Xml.val(sectPr, "type");
    s.type = type == null ? "nextPage" : type;
    Element pgNum = Xml.child(sectPr, "pgNumType");
    s.pageNumberStart = Xml.intAttr(pgNum, "start");
    String fmt = Xml.attr(pgNum, "fmt");
    s.pageNumberFormat = fmt == null ? "decimal" : fmt;
    for (Element ref : Xml.children(sectPr)) {
      boolean header = Xml.isW(ref, "headerReference");
      if (!header && !Xml.isW(ref, "footerReference")) {
        continue;
      }
      String kind = Xml.attr(ref, "type");
      PackagePart target = pkg.relatedById(pkg.mainPart(), Xml.relAttr(ref, "id"));
      if (kind == null || target == null) {
        continue;
      }
      (header ? s.headers : s.footers).put(kind, readStory(target));
    }
    return s;
  }

  private static float or(@Nullable Float value, float defaultValue) {
    return value == null ? defaultValue : value;
  }

  /** Reads a header or a footer part. */
  private List<Model.Block> readStory(PackagePart storyPart) throws IOException {
    PackagePart savedPart = part;
    Deque<FieldState> savedFields = fields;
    part = storyPart;
    fields = new ArrayDeque<>();
    try {
      return readBlocks(pkg.root(storyPart), null);
    } finally {
      part = savedPart;
      fields = savedFields;
    }
  }

  private List<Model.Block> readBlocks(@Nullable Element container, @Nullable TableContext tc)
      throws IOException {
    List<Model.Block> blocks = new ArrayList<>();
    readBlocks(container, tc, blocks);
    return blocks;
  }

  private void readBlocks(@Nullable Element container, @Nullable TableContext tc,
      List<Model.Block> out) throws IOException {
    for (Element e : Xml.children(container)) {
      if (Xml.isW(e, "p")) {
        out.add(readPara(e, tc));
      } else if (Xml.isW(e, "tbl")) {
        out.add(readTable(e));
      } else if (Xml.isW(e, "sdt")) {
        readBlocks(Xml.child(e, "sdtContent"), tc, out);
      } else if (Xml.isW(e, "customXml") || Xml.isW(e, "ins") || Xml.isW(e, "moveTo")
          || Xml.isW(e, "smartTag")) {
        readBlocks(e, tc, out);
      } else if (Xml.isW(e, "bookmarkStart")) {
        String name = Xml.attr(e, "name");
        if (name != null) {
          pendingBookmarks.add(name);
        }
      }
    }
  }

  // ---------------------------------------------------------------------------------------
  // paragraphs
  // ---------------------------------------------------------------------------------------

  private Model.Para readPara(Element p, @Nullable TableContext tc) throws IOException {
    Model.Para para = new Model.Para();
    Element ppr = Xml.child(p, "pPr");
    ParaProps direct = ParaProps.parse(ppr);
    String styleId = direct.styleId != null ? direct.styleId : styles.defaultParaStyle();
    StyleSheet.ParaBase base = styles.paraBase(styleId, tc == null ? null : tc.ppr());
    ParaProps props = base.props();

    Integer numId = direct.numId != null ? direct.numId : props.numId;
    Integer ilvlOrNull = direct.ilvl != null ? direct.ilvl : props.ilvl;
    int ilvl = ilvlOrNull == null ? 0 : ilvlOrNull;
    Numbering.Level level = numId == null || numId == 0 ? null : numbering.level(numId, ilvl);
    if (level != null && (direct.numId != null || base.numberingIndentApplies())) {
      props.applyIndent(level.ppr);
      props.applyTabs(level.ppr);
    }
    props.apply(direct);
    para.props = props;
    para.sectPr = Xml.child(ppr, "sectPr");

    RunProps runBase = styles.runBase(styleId, tc == null ? null : tc.rpr());
    RunProps mark = runBase.copy();
    RunProps markDirect = RunProps.parse(Xml.child(ppr, "rPr"), theme);
    if (markDirect.styleId != null) {
      mark.apply(styles.charStyle(markDirect.styleId));
    }
    mark.apply(markDirect);
    para.markProps = mark;

    // A paragraph whose mark is hidden or deleted still exists here, so it is numbered.
    if (level != null && numId != null) {
      Numbering.Label label = numbering.next(numId, ilvl);
      if (label != null) {
        para.label = label;
        RunProps labelProps = mark.copy();
        labelProps.underline = runBase.underline;
        labelProps.apply(label.rpr());
        para.labelProps = labelProps;
      }
    }

    for (String name : pendingBookmarks) {
      para.inlines.add(new Model.Bookmark(name));
    }
    pendingBookmarks.clear();
    readInlines(p, runBase, para.inlines);
    return para;
  }

  private void readInlines(Element parent, RunProps runBase, List<Model.Inline> out)
      throws IOException {
    for (Element e : Xml.children(parent)) {
      if (Xml.isW(e, "r")) {
        readRun(e, runBase, out);
      } else if (Xml.isW(e, "hyperlink")) {
        String saved = hyperlink;
        String anchor = Xml.attr(e, "anchor");
        hyperlink = anchor != null ? anchor : saved;
        readInlines(e, runBase, out);
        hyperlink = saved;
      } else if (Xml.isW(e, "sdt")) {
        Element content = Xml.child(e, "sdtContent");
        if (content != null) {
          readInlines(content, runBase, out);
        }
      } else if (Xml.isW(e, "ins") || Xml.isW(e, "moveTo") || Xml.isW(e, "smartTag")
          || Xml.isW(e, "customXml") || Xml.isW(e, "dir") || Xml.isW(e, "bdo")) {
        readInlines(e, runBase, out);
      } else if (Xml.isW(e, "fldSimple")) {
        FieldState fs = new FieldState();
        String instr = Xml.attr(e, "instr");
        fs.instruction.append(instr == null ? "" : instr);
        parseInstruction(fs);
        fs.inResult = true;
        fields.push(fs);
        readInlines(e, runBase, out);
        fields.pop();
        if (fs.dynamic) {
          out.add(toField(fs, runBase));
        }
      } else if (Xml.isW(e, "bookmarkStart")) {
        String name = Xml.attr(e, "name");
        if (name != null) {
          out.add(new Model.Bookmark(name));
        }
      }
    }
  }

  private void readRun(Element r, RunProps runBase, List<Model.Inline> out)
      throws IOException {
    RunProps rp = runBase.copy();
    RunProps direct = RunProps.parse(Xml.child(r, "rPr"), theme);
    if (direct.styleId != null) {
      rp.apply(styles.charStyle(direct.styleId));
    }
    rp.apply(direct);

    for (Element e : Xml.children(r)) {
      if (Xml.isW(e, "fldChar")) {
        fieldChar(Xml.attr(e, "fldCharType"), rp, out);
      } else if (Xml.isW(e, "instrText")) {
        FieldState top = fields.peek();
        if (top != null && !top.inResult) {
          top.instruction.append(e.getTextContent());
        }
      } else if (Xml.isW(e, "t")) {
        text(e.getTextContent(), rp, out);
      } else if (Xml.isW(e, "noBreakHyphen")) {
        text("‑", rp, out);
      } else if (Xml.isW(e, "sym")) {
        text(symbol(Xml.attr(e, "char")), rp, out);
      } else if (visible(rp)) {
        if (Xml.isW(e, "tab")) {
          out.add(new Model.Tab(rp, null));
        } else if (Xml.isW(e, "ptab")) {
          String alignment = Xml.attr(e, "alignment");
          out.add(new Model.Tab(rp, alignment == null ? "left" : alignment));
        } else if (Xml.isW(e, "br")) {
          String type = Xml.attr(e, "type");
          out.add(new Model.Break("page".equals(type) || "column".equals(type)));
        } else if (Xml.isW(e, "cr")) {
          out.add(new Model.Break(false));
        } else if (Xml.isW(e, "drawing")) {
          Model.Image image = readDrawing(e);
          if (image != null) {
            out.add(image);
          }
        } else if ("AlternateContent".equals(e.getLocalName()) || Xml.isW(e, "pict")
            || Xml.isW(e, "object")) {
          warnUnsupportedDrawing();
        }
      }
    }
  }

  /** Returns whether content is shown: not hidden, not in a field code or a replaced result. */
  private boolean visible(RunProps rp) {
    if (Boolean.TRUE.equals(rp.vanish)) {
      return false;
    }
    for (FieldState fs : fields) {
      if (!fs.inResult || fs.dynamic) {
        return false;
      }
    }
    return true;
  }

  private void text(@Nullable String raw, RunProps rp, List<Model.Inline> out) {
    if (raw == null || raw.isEmpty()) {
      return;
    }
    FieldState top = fields.peek();
    if (top != null && !top.inResult) {
      return;
    }
    for (FieldState fs : fields) {
      if (fs.dynamic && fs.inResult) {
        fs.fallback.append(raw);
        if (fs.rpr == null) {
          fs.rpr = rp;
        }
        return;
      }
    }
    if (Boolean.TRUE.equals(rp.vanish)) {
      return;
    }
    StringBuilder sb = new StringBuilder();
    for (int i = 0; i < raw.length(); i++) {
      char c = raw.charAt(i);
      if (c == '\t') {
        flush(sb, rp, out);
        out.add(new Model.Tab(rp, null));
      } else if (c >= 0x20 && c != 0x00AD && c != 0xFEFF && !(c >= 0x200B && c <= 0x200F)
          && !(c >= 0x202A && c <= 0x202E) && !(c >= 0x2060 && c <= 0x2069)) {
        sb.append(c);
      }
    }
    flush(sb, rp, out);
  }

  private void flush(StringBuilder sb, RunProps rp, List<Model.Inline> out) {
    if (sb.length() > 0) {
      out.add(new Model.Text(sb.toString(), rp, currentLink()));
      sb.setLength(0);
    }
  }

  private @Nullable String currentLink() {
    if (hyperlink != null) {
      return hyperlink;
    }
    for (FieldState fs : fields) {
      if (fs.link != null) {
        return fs.link;
      }
    }
    return null;
  }

  private static String symbol(@Nullable String hex) {
    if (hex == null) {
      return "";
    }
    try {
      int code = Integer.parseInt(hex, 16);
      // Symbol fonts are addressed in the private use area: only common bullets are mapped.
      return switch (code) {
        case 0xF0B7 -> "•";
        case 0xF0A7 -> "▪";
        case 0xF06C -> "●";
        case 0xF06E -> "■";
        case 0xF0FC -> "✓";
        case 0xF0A8 -> "◆";
        default -> code >= 0xF000 && code <= 0xF0FF ? "" : new String(Character.toChars(code));
      };
    } catch (IllegalArgumentException e) {
      return "";
    }
  }

  // ---------------------------------------------------------------------------------------
  // fields
  // ---------------------------------------------------------------------------------------

  private void fieldChar(@Nullable String type, RunProps rp, List<Model.Inline> out) {
    if ("begin".equals(type)) {
      fields.push(new FieldState());
    } else if ("separate".equals(type)) {
      FieldState top = fields.peek();
      if (top != null) {
        parseInstruction(top);
        top.inResult = true;
      }
    } else if ("end".equals(type) && !fields.isEmpty()) {
      FieldState fs = fields.pop();
      parseInstruction(fs);
      if (fs.dynamic && visibleOutside()) {
        out.add(toField(fs, rp));
      }
    }
  }

  /** Returns whether the place just outside the popped field is shown. */
  private boolean visibleOutside() {
    for (FieldState fs : fields) {
      if (!fs.inResult || fs.dynamic) {
        return false;
      }
    }
    return true;
  }

  private Model.Field toField(FieldState fs, RunProps fallbackProps) {
    RunProps fieldProps = fs.rpr;
    RunProps rp = fieldProps != null ? fieldProps : fallbackProps;
    return new Model.Field(fs.kind, fs.arg, fs.format, rp, fs.fallback.toString().trim(),
        currentLink());
  }

  private static void parseInstruction(FieldState fs) {
    if (fs.parsed) {
      return;
    }
    fs.parsed = true;
    List<String> tokens = tokenize(fs.instruction.toString());
    if (tokens.isEmpty()) {
      return;
    }
    fs.kind = tokens.get(0).toUpperCase(Locale.ROOT);
    fs.dynamic = DYNAMIC_FIELDS.contains(fs.kind);
    for (int i = 1; i < tokens.size(); i++) {
      String t = tokens.get(i);
      if (t.startsWith("\\")) {
        boolean hasNext = i + 1 < tokens.size();
        if ("\\*".equals(t) && hasNext) {
          fs.format = pageFormat(tokens.get(++i));
        } else if ("\\l".equals(t) && hasNext && "HYPERLINK".equals(fs.kind)) {
          fs.link = tokens.get(++i);
        }
      } else if (fs.arg.isEmpty()) {
        fs.arg = t;
      }
    }
    if ("PAGEREF".equals(fs.kind) && !fs.arg.isEmpty()) {
      fs.link = fs.arg;
    }
  }

  private static String pageFormat(String switchValue) {
    return switch (switchValue) {
      case "roman" -> "lowerRoman";
      case "ROMAN", "Roman" -> "upperRoman";
      case "alphabetic" -> "lowerLetter";
      case "ALPHABETIC", "Alphabetic" -> "upperLetter";
      case "Arabic", "arabic", "ARABIC" -> "decimal";
      case "DBCHAR", "dbchar" -> "decimalFullWidth";
      case "CIRCLENUM" -> "decimalEnclosedCircle";
      default -> "";
    };
  }

  private static List<String> tokenize(String instruction) {
    List<String> tokens = new ArrayList<>();
    StringBuilder sb = new StringBuilder();
    boolean quoted = false;
    for (int i = 0; i < instruction.length(); i++) {
      char c = instruction.charAt(i);
      if (c == '"') {
        quoted = !quoted;
      } else if (Character.isWhitespace(c) && !quoted) {
        if (sb.length() > 0) {
          tokens.add(sb.toString());
          sb.setLength(0);
        }
      } else {
        sb.append(c);
      }
    }
    if (sb.length() > 0) {
      tokens.add(sb.toString());
    }
    return tokens;
  }

  // ---------------------------------------------------------------------------------------
  // pictures
  // ---------------------------------------------------------------------------------------

  private Model.@Nullable Image readDrawing(Element drawing) throws IOException {
    Element inline = Xml.childAnyNs(drawing, "inline");
    Element anchorEl = Xml.childAnyNs(drawing, "anchor");
    Element container = inline != null ? inline : anchorEl;
    Element blip = Xml.descendant(container, "blip");
    Element extent = Xml.childAnyNs(container, "extent");
    String relId = Xml.relAttr(blip, "embed");
    Integer cx = Xml.intAttr(extent, "cx");
    Integer cy = Xml.intAttr(extent, "cy");
    PackagePart imagePart = pkg.relatedById(part, relId);
    if (container == null || relId == null || imagePart == null || cx == null || cy == null) {
      warnUnsupportedDrawing();
      return null;
    }
    Map<String, Model.ImageData> cache =
        Objects.requireNonNull(images.computeIfAbsent(part, k -> new HashMap<>()));
    Model.ImageData data = cache.get(relId);
    if (data == null) {
      data = new Model.ImageData(pkg.bytes(imagePart));
      cache.put(relId, data);
    }
    Model.Anchor anchor = null;
    if (anchorEl != null) {
      Element h = Xml.childAnyNs(anchorEl, "positionH");
      Element v = Xml.childAnyNs(anchorEl, "positionV");
      String relH = Xml.attr(h, "relativeFrom");
      String relV = Xml.attr(v, "relativeFrom");
      anchor = new Model.Anchor(relH == null ? "column" : relH, offset(h), align(h),
          relV == null ? "paragraph" : relV, offset(v), align(v));
    }
    return new Model.Image(data, cx / EMU_PER_POINT, cy / EMU_PER_POINT, anchor);
  }

  private static float offset(@Nullable Element position) {
    Element e = Xml.childAnyNs(position, "posOffset");
    if (e == null) {
      return 0f;
    }
    try {
      return Float.parseFloat(e.getTextContent().trim()) / EMU_PER_POINT;
    } catch (NumberFormatException ex) {
      return 0f;
    }
  }

  private static @Nullable String align(@Nullable Element position) {
    Element e = Xml.childAnyNs(position, "align");
    return e == null ? null : e.getTextContent().trim();
  }

  private void warnUnsupportedDrawing() {
    if (!warnedUnsupportedDrawing) {
      warnedUnsupportedDrawing = true;
      detailLog.warn("The Word file has drawings other than pictures (shapes, text boxes, "
          + "charts etc.). They are not supported and are not output to the PDF file.");
    }
  }

  // ---------------------------------------------------------------------------------------
  // tables
  // ---------------------------------------------------------------------------------------

  private Model.Table readTable(Element tbl) throws IOException {
    Model.Table table = new Model.Table();
    Element tblPr = Xml.child(tbl, "tblPr");
    StyleSheet.TableStyle style = styles.tableStyle(Xml.val(tblPr, "tblStyle"));
    table.borders.putAll(style.borders);
    table.borders.putAll(Border.parseAll(Xml.child(tblPr, "tblBorders")));
    Float[] margins = style.cellMargins.clone();
    StyleSheet.applyCellMargins(margins, Xml.child(tblPr, "tblCellMar"));
    for (int i = 0; i < 4; i++) {
      Float m = margins[i];
      if (m != null) {
        table.cellMargins[i] = m;
      }
    }
    Element ind = Xml.child(tblPr, "tblInd");
    if (!"pct".equals(Xml.attr(ind, "type"))) {
      table.indent = or(Xml.twips(ind, "w"), 0f);
    }
    table.fixedLayout = "fixed".equals(Xml.attr(Xml.child(tblPr, "tblLayout"), "type"));
    String jc = Xml.val(tblPr, "jc");
    table.jc = jc == null || "start".equals(jc) ? "left" : "end".equals(jc) ? "right" : jc;
    Element tblW = Xml.child(tblPr, "tblW");
    String w = Xml.attr(tblW, "w");
    if (w != null && w.endsWith("%")) {
      table.widthRatio = parseFloat(w.substring(0, w.length() - 1)) / 100f;
    } else if (w != null && "pct".equals(Xml.attr(tblW, "type"))) {
      table.widthRatio = parseFloat(w) / 5000f;
    }
    List<Element> cols = Xml.children(Xml.child(tbl, "tblGrid"), "gridCol");
    table.grid = new float[cols.size()];
    for (int i = 0; i < cols.size(); i++) {
      table.grid[i] = or(Xml.twips(cols.get(i), "w"), 0f);
    }

    Element look = Xml.child(tblPr, "tblLook");
    int lookBits = 0;
    String lookVal = Xml.attr(look, "val");
    if (lookVal != null) {
      try {
        lookBits = Integer.parseInt(lookVal, 16);
      } catch (NumberFormatException e) {
        lookBits = 0;
      }
    }
    final boolean firstRow = lookFlag(look, "firstRow", lookBits, 0x0020);
    final boolean lastRow = lookFlag(look, "lastRow", lookBits, 0x0040);
    final boolean firstCol = lookFlag(look, "firstColumn", lookBits, 0x0080);
    final boolean lastCol = lookFlag(look, "lastColumn", lookBits, 0x0100);
    final boolean noHorzBand = lookFlag(look, "noHBand", lookBits, 0x0200);

    List<Element> rows = new ArrayList<>();
    collectRows(tbl, rows);
    for (int r = 0; r < rows.size(); r++) {
      Element tr = rows.get(r);
      Model.Row row = new Model.Row();
      Element trPr = Xml.child(tr, "trPr");
      Element height = Xml.child(trPr, "trHeight");
      row.height = or(Xml.twips(height, "val"), 0f);
      row.exactHeight = "exact".equals(Xml.attr(height, "hRule"));
      row.cantSplit = Boolean.TRUE.equals(Xml.onOff(trPr, "cantSplit"));
      row.header = Boolean.TRUE.equals(Xml.onOff(trPr, "tblHeader"));
      Integer gridBefore = Xml.intAttr(Xml.child(trPr, "gridBefore"), "val");
      row.gridBefore = gridBefore == null ? 0 : gridBefore;

      List<Element> cells = new ArrayList<>();
      collectCells(tr, cells);
      for (int c = 0; c < cells.size(); c++) {
        List<StyleSheet.TablePart> parts = new ArrayList<>();
        parts.add(style.whole);
        int band = r - (firstRow ? 1 : 0);
        if (!noHorzBand && band >= 0) {
          addPart(parts, style, band % 2 == 0 ? "band1Horz" : "band2Horz");
        }
        if (firstCol && c == 0) {
          addPart(parts, style, "firstCol");
        }
        if (lastCol && c == cells.size() - 1) {
          addPart(parts, style, "lastCol");
        }
        if (firstRow && r == 0) {
          addPart(parts, style, "firstRow");
        }
        if (lastRow && r == rows.size() - 1 && r > 0) {
          addPart(parts, style, "lastRow");
        }
        row.cells.add(readCell(cells.get(c), parts));
      }
      table.rows.add(row);
    }
    return table;
  }

  private static void addPart(List<StyleSheet.TablePart> parts, StyleSheet.TableStyle style,
      String type) {
    StyleSheet.TablePart part = style.conditions.get(type);
    if (part != null) {
      parts.add(part);
    }
  }

  private static boolean lookFlag(@Nullable Element look, String attr, int bits, int mask) {
    String v = Xml.attr(look, attr);
    if (v != null) {
      return v.equals("1") || v.equals("true") || v.equals("on");
    }
    return (bits & mask) != 0;
  }

  private static float parseFloat(String s) {
    try {
      return Float.parseFloat(s.trim());
    } catch (NumberFormatException e) {
      return 0f;
    }
  }

  private static void collectRows(Element parent, List<Element> out) {
    for (Element e : Xml.children(parent)) {
      if (Xml.isW(e, "tr")) {
        out.add(e);
      } else if (Xml.isW(e, "sdt")) {
        Element content = Xml.child(e, "sdtContent");
        if (content != null) {
          collectRows(content, out);
        }
      } else if (Xml.isW(e, "ins") || Xml.isW(e, "customXml")) {
        collectRows(e, out);
      }
    }
  }

  private static void collectCells(Element parent, List<Element> out) {
    for (Element e : Xml.children(parent)) {
      if (Xml.isW(e, "tc")) {
        out.add(e);
      } else if (Xml.isW(e, "sdt")) {
        Element content = Xml.child(e, "sdtContent");
        if (content != null) {
          collectCells(content, out);
        }
      } else if (Xml.isW(e, "ins") || Xml.isW(e, "customXml")) {
        collectCells(e, out);
      }
    }
  }

  private Model.Cell readCell(Element tc, List<StyleSheet.TablePart> parts) throws IOException {
    Model.Cell cell = new Model.Cell();
    ParaProps ppr = new ParaProps();
    RunProps rpr = new RunProps();
    for (StyleSheet.TablePart part : parts) {
      ppr.apply(part.ppr);
      rpr.apply(part.rpr);
      cell.borders.putAll(part.cellBorders);
      if (part.shade != null) {
        cell.shade = part.shade;
      }
    }
    Element tcPr = Xml.child(tc, "tcPr");
    Integer span = Xml.intAttr(Xml.child(tcPr, "gridSpan"), "val");
    cell.span = span == null || span < 1 ? 1 : span;
    Element vmerge = Xml.child(tcPr, "vMerge");
    if (vmerge != null) {
      cell.vmerge = "restart".equals(Xml.attr(vmerge, "val")) ? 1 : 2;
    }
    cell.borders.putAll(Border.parseAll(Xml.child(tcPr, "tcBorders")));
    Element shd = Xml.child(tcPr, "shd");
    if (shd != null) {
      Integer fill = Xml.color(Xml.attr(shd, "fill"));
      cell.shade = fill != null ? fill : RunProps.NO_COLOR;
    }
    String valign = Xml.val(tcPr, "vAlign");
    cell.valign = valign == null ? "top" : valign;
    StyleSheet.applyCellMargins(cell.margins, Xml.child(tcPr, "tcMar"));
    cell.blocks = readBlocks(tc, new TableContext(ppr, rpr));
    if (cell.blocks.isEmpty()) {
      cell.blocks.add(new Model.Para());
    }
    return cell;
  }
}
