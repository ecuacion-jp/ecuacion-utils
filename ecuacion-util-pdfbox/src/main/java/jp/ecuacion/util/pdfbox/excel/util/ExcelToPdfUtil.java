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
package jp.ecuacion.util.pdfbox.excel.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdfbox.excel.exception.PdfGenerateException;
import jp.ecuacion.util.pdfbox.excel.internal.FontManager;
import jp.ecuacion.util.pdfbox.excel.internal.SheetRenderer;
import jp.ecuacion.util.pdfbox.excel.options.PdfGenerateOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jspecify.annotations.Nullable;

/**
 * Provides utility methods for generating PDF files from Excel files.
 *
 * <p>The output PDF reflects the print area of each specified sheet, preserving
 * cell values, text styles, background colors, borders, and merged cells
 * as closely as possible to the Excel appearance.</p>
 *
 * <p>Font files (Noto Sans JP) must be present at
 * {@code src/main/resources/fonts/NotoSansJP/NotoSansJP-Regular.ttf} and
 * {@code src/main/resources/fonts/NotoSansJP/NotoSansJP-Bold.ttf}.</p>
 */
public class ExcelToPdfUtil {

  private ExcelToPdfUtil() {}

  /**
   * Generates a PDF file from the specified sheets of an Excel file.
   *
   * <p>Sheets are rendered in the order given by {@code sheetNames}.
   * Each sheet's print area (and page breaks) determines how many PDF pages are produced.</p>
   *
   * @param excelPath path to the source Excel file
   * @param sheetNames list of sheet names to include in the PDF, in order
   * @param outputPath path to the output PDF file
   * @param options optional parameters (passwords, etc.); may be {@code null}
   * @throws PdfGenerateException if an error occurs during PDF generation
   */
  public static void generate(Path excelPath, List<String> sheetNames,
      Path outputPath, @Nullable PdfGenerateOptions options) throws PdfGenerateException {

    String excelPassword = (options != null) ? options.getExcelPassword() : null;
    String pdfPassword = (options != null) ? options.getPdfPassword() : null;

    File excelFile = excelPath.toFile();

    try (Workbook workbook = openWorkbook(excelFile, excelPassword);
        PDDocument document = new PDDocument()) {

      FontManager fontManager = new FontManager(document);
      SheetRenderer renderer = new SheetRenderer(document, fontManager);

      for (String sheetName : sheetNames) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
          throw new PdfGenerateException("Sheet not found: '" + sheetName + "'");
        }
        renderer.render(workbook, sheetIndex);
      }

      if (pdfPassword != null) {
        AccessPermission ap = new AccessPermission();
        StandardProtectionPolicy policy =
            new StandardProtectionPolicy(pdfPassword, pdfPassword, ap);
        document.protect(policy);
      }

      document.save(outputPath.toFile());

    } catch (PdfGenerateException e) {
      throw e;
    } catch (IOException e) {
      throw new PdfGenerateException("Failed to generate PDF from '" + excelPath + "'", e);
    }
  }

  private static Workbook openWorkbook(File file, @Nullable String password) throws IOException {
    if (password != null) {
      return WorkbookFactory.create(file, password, true);
    }
    return WorkbookFactory.create(file, null, true);
  }
}
