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

/**
 * Provides PDF generation from Word documents using Apache PDFBox.
 */
module jp.ecuacion.util.pdf.word.document {
  exports jp.ecuacion.util.pdf.word.document.options;
  exports jp.ecuacion.util.pdf.word.document.util;

  requires jp.ecuacion.lib.core;
  requires transitive jp.ecuacion.util.pdf.common;

  requires org.apache.poi.poi;
  requires org.apache.poi.ooxml;
  requires org.apache.pdfbox;
  requires org.apache.pdfbox.io;
  requires org.apache.fontbox;
  requires java.xml;
  // Pictures are decoded with ImageIO (through PDFBox).
  requires java.desktop;
}
