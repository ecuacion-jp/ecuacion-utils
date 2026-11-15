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
import org.w3c.dom.Element;

/** Holds the document model read from a .docx file, with the properties already resolved. */
final class Model {

  private Model() {}

  /** A block-level element: a paragraph or a table. */
  interface Block {
  }

  /** An inline element of a paragraph. */
  interface Inline {
  }

  /** A paragraph. */
  static final class Para implements Block {
    ParaProps props = new ParaProps();
    /** The run properties of the paragraph mark (used for empty lines and number labels). */
    RunProps markProps = new RunProps();
    final List<Inline> inlines = new ArrayList<>();
    Numbering.@Nullable Label label;
    @Nullable RunProps labelProps;
    /** The section properties, when this paragraph is the last one of a section. */
    @Nullable Element sectPr;
  }

  /** A piece of text with the same formatting. */
  record Text(String text, RunProps rpr, @Nullable String link) implements Inline {
  }

  /** A tab. {@code absolute} is center / right for a positional tab ({@code w:ptab}). */
  record Tab(RunProps rpr, @Nullable String absolute) implements Inline {
  }

  /** A break. */
  record Break(boolean page) implements Inline {
  }

  /** A picture. The size is in points. */
  record Image(ImageData data, float width, float height, @Nullable Anchor anchor)
      implements Inline {
  }

  /** The bytes of a picture (an object, so that the same picture is embedded only once). */
  static final class ImageData {
    private final byte[] bytes;

    ImageData(byte[] bytes) {
      this.bytes = bytes.clone();
    }

    byte[] bytes() {
      return bytes.clone();
    }
  }

  /** The position of a floating picture. Offsets are in points. */
  record Anchor(String relH, float offsetX, @Nullable String alignH, String relV,
      float offsetY, @Nullable String alignV) {
  }

  /** A field whose value is decided at layout time: PAGE, NUMPAGES, SECTIONPAGES, PAGEREF. */
  record Field(String kind, String arg, String format, RunProps rpr, String fallback,
      @Nullable String link) implements Inline {
  }

  /** The start of a bookmark (the target of PAGEREF fields and internal links). */
  record Bookmark(String name) implements Inline {
  }

  /** A table. Lengths are in points. */
  static final class Table implements Block {
    float[] grid = new float[0];
    final List<Row> rows = new ArrayList<>();
    float indent;
    String jc = "left";
    /** The width as a ratio of the available width, when specified in percent. */
    @Nullable Float widthRatio;
    Map<String, Border> borders = new HashMap<>();
    float[] cellMargins = {5.4f, 0f, 5.4f, 0f};
    /** Whether the column widths are fixed (not adjusted to the content). */
    boolean fixedLayout;
  }

  /** A table row. */
  static final class Row {
    final List<Cell> cells = new ArrayList<>();
    float height;
    boolean exactHeight;
    boolean cantSplit;
    boolean header;
    int gridBefore;
  }

  /** A table cell. */
  static final class Cell {
    int span = 1;
    /** 0: not merged, 1: the first cell of a vertical merge, 2: a continued cell. */
    int vmerge;
    List<Block> blocks = new ArrayList<>();
    Map<String, Border> borders = new HashMap<>();
    @Nullable Integer shade;
    String valign = "top";
    @Nullable Float[] margins = new Float[4];
  }

  /** A section: page settings, headers / footers and the content. Lengths are in points. */
  static final class Section {
    float pageWidth = 595.3f;
    float pageHeight = 841.9f;
    float marginTop = 72f;
    float marginBottom = 72f;
    float marginLeft = 72f;
    float marginRight = 72f;
    float headerDistance = 36f;
    float footerDistance = 36f;
    /** The line pitch of the document grid, or 0 when lines are not snapped to a grid. */
    float linePitch;
    /** The extra advance per East Asian character of the character grid. */
    float charDelta;
    boolean titlePage;
    String type = "nextPage";
    @Nullable Integer pageNumberStart;
    String pageNumberFormat = "decimal";
    /** default / first / even -> blocks. */
    Map<String, List<Block>> headers = new HashMap<>();
    Map<String, List<Block>> footers = new HashMap<>();
    final List<Block> blocks = new ArrayList<>();
  }

  /** Document-wide settings. */
  static final class Settings {
    float defaultTab = 36f;
    boolean evenAndOddHeaders;
    boolean compressPunctuation;
    int compatibilityMode = 12;
    /** Whether the lines in table cells snap to the document grid. */
    boolean snapToGridInCell;
  }

  /** A document. */
  static final class Document {
    final List<Section> sections = new ArrayList<>();
    Settings settings = new Settings();
    /** The run properties of the default paragraph style. */
    RunProps defaultRun = new RunProps();
  }
}
