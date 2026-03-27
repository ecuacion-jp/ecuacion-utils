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
 * Provides PDF generation utilities using Apache PDFBox.
 */
module jp.ecuacion.util.pdfbox {
  exports jp.ecuacion.util.pdfbox.excel.exception;
  exports jp.ecuacion.util.pdfbox.excel.options;
  exports jp.ecuacion.util.pdfbox.excel.util;

  provides jp.ecuacion.lib.core.spi.MessagesUtilPdfboxProvider
      with jp.ecuacion.util.pdfbox.spi.impl.internal.MessagesUtilPdfboxProviderImpl;

  requires jakarta.annotation;
  requires java.desktop;
  requires jp.ecuacion.lib.core;

  requires transitive org.apache.poi.poi;
  requires org.apache.poi.ooxml;
  requires org.apache.pdfbox;
  requires org.apache.commons.lang3;
  requires java.desktop;
  requires org.apache.fontbox;
}
