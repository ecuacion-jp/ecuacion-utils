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
import java.util.List;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;

/**
 * Provides helpers to read WordprocessingML with DOM.
 *
 * <p>Elements are matched by their local names. Both the transitional and the strict
 * namespaces of WordprocessingML are accepted.</p>
 */
final class Xml {

  private static final String REL_NS_SUFFIX = "/relationships";

  private Xml() {}

  /** Returns whether the node is the WordprocessingML element with the local name. */
  static boolean isW(Node node, String local) {
    if (!(node instanceof Element) || !local.equals(node.getLocalName())) {
      return false;
    }
    String ns = node.getNamespaceURI();
    // transitional: .../wordprocessingml/2006/main, strict: .../ooxml/wordprocessingml/main
    return ns != null && ns.contains("/wordprocessingml/") && ns.endsWith("/main");
  }

  /** Returns all the child elements. */
  static List<Element> children(@Nullable Element parent) {
    List<Element> list = new ArrayList<>();
    if (parent == null) {
      return list;
    }
    for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
      if (n instanceof Element e) {
        list.add(e);
      }
    }
    return list;
  }

  /** Returns the WordprocessingML child elements with the local name. */
  static List<Element> children(@Nullable Element parent, String local) {
    List<Element> list = new ArrayList<>();
    for (Element e : children(parent)) {
      if (isW(e, local)) {
        list.add(e);
      }
    }
    return list;
  }

  /** Returns the first WordprocessingML child element with the local name. */
  static @Nullable Element child(@Nullable Element parent, String local) {
    if (parent == null) {
      return null;
    }
    for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
      if (isW(n, local)) {
        return (Element) n;
      }
    }
    return null;
  }

  /** Returns the first child element with the local name, in any namespace. */
  static @Nullable Element childAnyNs(@Nullable Element parent, String local) {
    if (parent == null) {
      return null;
    }
    for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
      if (n instanceof Element e && local.equals(e.getLocalName())) {
        return e;
      }
    }
    return null;
  }

  /** Returns the first descendant element with the local name, in any namespace. */
  static @Nullable Element descendant(@Nullable Element parent, String local) {
    if (parent == null) {
      return null;
    }
    for (Node n = parent.getFirstChild(); n != null; n = n.getNextSibling()) {
      if (n instanceof Element e) {
        if (local.equals(e.getLocalName())) {
          return e;
        }
        Element found = descendant(e, local);
        if (found != null) {
          return found;
        }
      }
    }
    return null;
  }

  /** Returns the attribute with the local name (other than relationship attributes). */
  static @Nullable String attr(@Nullable Element e, String local) {
    return attribute(e, local, false);
  }

  /** Returns the relationship attribute (e.g. {@code r:id}) with the local name. */
  static @Nullable String relAttr(@Nullable Element e, String local) {
    return attribute(e, local, true);
  }

  private static @Nullable String attribute(@Nullable Element e, String local, boolean rel) {
    if (e == null) {
      return null;
    }
    NamedNodeMap attrs = e.getAttributes();
    for (int i = 0; i < attrs.getLength(); i++) {
      Node a = attrs.item(i);
      String name = a.getLocalName() != null ? a.getLocalName() : a.getNodeName();
      if (!local.equals(name)) {
        continue;
      }
      String ns = a.getNamespaceURI();
      boolean isRel = ns != null && ns.endsWith(REL_NS_SUFFIX);
      if (isRel == rel) {
        return a.getNodeValue();
      }
    }
    return null;
  }

  /** Returns the {@code w:val} of the child element. */
  static @Nullable String val(@Nullable Element parent, String childLocal) {
    return attr(child(parent, childLocal), "val");
  }

  /** Returns the attribute as an integer. */
  static @Nullable Integer intAttr(@Nullable Element e, String local) {
    String v = attr(e, local);
    if (v == null || v.isBlank()) {
      return null;
    }
    try {
      return (int) Math.round(Double.parseDouble(v.trim()));
    } catch (NumberFormatException ex) {
      return null;
    }
  }

  /** Returns the attribute in twips (1/20 pt) as points. */
  static @Nullable Float twips(@Nullable Element e, String local) {
    String v = attr(e, local);
    if (v == null || v.isBlank()) {
      return null;
    }
    try {
      return (float) (Double.parseDouble(v.trim()) / 20.0);
    } catch (NumberFormatException ex) {
      return null;
    }
  }

  /** Returns the on/off value of the child element, or {@code null} if it does not exist. */
  static @Nullable Boolean onOff(@Nullable Element parent, String childLocal) {
    Element e = child(parent, childLocal);
    if (e == null) {
      return null;
    }
    String v = attr(e, "val");
    return v == null || !(v.equals("0") || v.equals("false") || v.equals("off"));
  }

  /** Parses a color like {@code RRGGBB}. Returns {@code null} for {@code auto} and others. */
  static @Nullable Integer color(@Nullable String v) {
    if (v == null || v.length() != 6) {
      return null;
    }
    try {
      return Integer.parseInt(v, 16);
    } catch (NumberFormatException ex) {
      return null;
    }
  }
}
