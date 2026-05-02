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
module jp.ecuacion.util.pdf.excel.report {
  exports jp.ecuacion.util.pdf.excel.report.exception;
  exports jp.ecuacion.util.pdf.excel.report.options;
  exports jp.ecuacion.util.pdf.excel.report.util;

  provides jp.ecuacion.lib.core.spi.MessagesUtilExcelReportToPdfProvider
      with jp.ecuacion.util.pdf.excel.report.spi.impl.internal.MessagesUtilExcelReportToPdfProviderImpl;

  requires jakarta.annotation;
  requires jp.ecuacion.lib.core;

  requires transitive org.apache.poi.poi;
  requires org.apache.poi.ooxml;
  requires org.apache.pdfbox;
  requires org.apache.commons.lang3;
  requires org.apache.fontbox;
  requires java.desktop;
}
