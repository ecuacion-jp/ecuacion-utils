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
 * Provides the features shared by the ecuacion-util modules which generate PDF files using
 * Apache PDFBox (font resolution and the exceptions related to it).
 */
// "module": the internal package is exported only to the excel / word modules, which are
// built after this module (javac warns that they are not found).
@SuppressWarnings("module")
module jp.ecuacion.util.pdf.common {
  exports jp.ecuacion.util.pdf.common.exception;
  exports jp.ecuacion.util.pdf.common.internal
      to jp.ecuacion.util.pdf.excel.report, jp.ecuacion.util.pdf.word.document;

  provides jp.ecuacion.lib.core.spi.MessagesUtilPdfCommonProvider
      with jp.ecuacion.util.pdf.common.spi.impl.internal.MessagesUtilPdfCommonProviderImpl;

  requires jp.ecuacion.lib.core;

  requires org.apache.pdfbox;
  requires org.apache.pdfbox.io;
  requires org.apache.fontbox;
}
