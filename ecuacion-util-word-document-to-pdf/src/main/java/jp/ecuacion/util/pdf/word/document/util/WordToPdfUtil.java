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
import java.io.UncheckedIOException;
import java.nio.file.Path;
import jp.ecuacion.util.pdf.common.exception.PdfGenerateException;
import jp.ecuacion.util.pdf.word.document.internal.WordRenderer;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;

/**
 * Provides utility methods for generating PDF files from Word files.
 *
 * <p>It is meant for text-centered documents such as contracts and terms of service.
 * Paragraph and character formatting, numbered headings and lists, tables, headers / footers,
 * page numbers, tables of contents and pictures are output as closely as possible to the
 * appearance in Word. Shapes, text boxes, charts, footnotes and multiple columns are not
 * supported.</p>
 *
 * <p>Fonts must be specified via {@link PdfGenerateOptions#builderForSystemFonts()} or
 * {@link PdfGenerateOptions#builderForExplicitFont(java.nio.file.Path)}.</p>
 */
public class WordToPdfUtil {

  private WordToPdfUtil() {}

  /**
   * Generates a PDF file from a Word file: {@code .docx}, and also {@code .docm}, {@code .dotx}
   * and {@code .dotm} (macros are not run).
   *
   * <p>The page numbers in a table of contents and in cross-references are recalculated from
   * the layout of the PDF file, so they are correct even when the fields in the Word file are
   * not updated.</p>
   *
   * <p>When {@link PdfGenerateOptions#getPdfPassword()} is set, the output PDF is encrypted
   * with 256-bit AES.</p>
   *
   * @param wordPath path to the source Word file
   * @param outputPath path to the output PDF file
   * @param options parameters including font paths and optional passwords
   * @throws PdfGenerateException if the Word file or the font environment prevents PDF
   *     generation (e.g. the file is not a Word file or its password is wrong, a font is
   *     missing, or a character cannot be rendered)
   * @throws UncheckedIOException if a technical failure occurs while reading the Word file or
   *     writing the PDF file
   */
  public static void generate(Path wordPath, Path outputPath, PdfGenerateOptions options) {
    String pdfPassword = options.getPdfPassword();

    try (PDDocument document = new PDDocument()) {
      WordRenderer.render(wordPath, document, options);

      if (pdfPassword != null) {
        AccessPermission ap = new AccessPermission();
        String pdfOwnerPassword =
            options.getPdfOwnerPassword() != null ? options.getPdfOwnerPassword() : pdfPassword;
        StandardProtectionPolicy policy =
            new StandardProtectionPolicy(pdfOwnerPassword, pdfPassword, ap);
        // PDFBox defaults to 40-bit RC4, which is trivially breakable.
        policy.setEncryptionKeyLength(256);
        policy.setPreferAES(true);
        document.protect(policy);
      }

      document.save(outputPath.toFile());

    } catch (IOException e) {
      throw new UncheckedIOException("Failed to generate PDF from '" + wordPath + "'", e);
    }
  }
}
