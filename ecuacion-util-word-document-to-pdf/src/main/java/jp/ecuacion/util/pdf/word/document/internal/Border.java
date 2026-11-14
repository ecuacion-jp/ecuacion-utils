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
import java.util.Map;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;

/**
 * Holds a border of a paragraph or a table cell.
 *
 * @param style the border style ({@code none} for an explicitly removed border)
 * @param width the line width in points
 * @param space the space between the border and the content in points
 * @param color the RGB color
 */
record Border(String style, float width, float space, int color) {

  /** Returns the room the border takes: a double border is two lines and the gap of them. */
  float room() {
    return "double".equals(style) ? width * 3 : width;
  }

  boolean visible() {
    return !"none".equals(style) && !"nil".equals(style) && width > 0;
  }

  static Border parse(Element e) {
    String val = Xml.attr(e, "val");
    Integer sz = Xml.intAttr(e, "sz");
    Integer space = Xml.intAttr(e, "space");
    Integer color = Xml.color(Xml.attr(e, "color"));
    String style = val == null ? "single" : val;
    float width = sz == null ? 0.5f : sz / 8f;
    if ("nil".equals(style)) {
      style = "none";
    }
    return new Border(style, width, space == null ? 0f : space, color == null ? 0 : color);
  }

  /** Parses the border elements under a container like {@code pBdr} or {@code tcBorders}. */
  static Map<String, Border> parseAll(@Nullable Element container) {
    Map<String, Border> map = new HashMap<>();
    for (Element e : Xml.children(container)) {
      String name = e.getLocalName();
      if ("start".equals(name)) {
        name = "left";
      } else if ("end".equals(name)) {
        name = "right";
      }
      map.put(name, parse(e));
    }
    return map;
  }
}
