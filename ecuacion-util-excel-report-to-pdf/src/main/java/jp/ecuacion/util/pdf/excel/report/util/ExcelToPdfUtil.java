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
package jp.ecuacion.util.pdf.excel.report.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import jp.ecuacion.lib.core.util.LocaleUtil;
import jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException;
import jp.ecuacion.util.pdf.excel.report.internal.FontManager;
import jp.ecuacion.util.pdf.excel.report.internal.SheetRenderer;
import jp.ecuacion.util.pdf.excel.report.internal.SystemFontLocator;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;

/**
 * Provides utility methods for generating PDF files from Excel files.
 *
 * <p>The output PDF reflects the print area of each specified sheet, preserving
 * cell values, text styles, background colors, borders, and merged cells
 * as closely as possible to the Excel appearance.</p>
 *
 * <p>Font paths must be specified via {@link PdfGenerateOptions#builder()}.</p>
 */
public class ExcelToPdfUtil {

  private ExcelToPdfUtil() {}

  /**
   * Generates a PDF file from the specified sheets of an Excel file.
   *
   * <p>Sheets are rendered in the order given by {@code sheetNames}.
   * Each sheet's print area (and page breaks) determines how many PDF pages are produced.</p>
   *
   * <p>When {@link PdfGenerateOptions#getPdfPassword()} is set, the output PDF is encrypted
   * with 256-bit AES.</p>
   *
   * @param excelPath path to the source Excel file
   * @param sheetNames list of sheet names to include in the PDF, in order
   * @param outputPath path to the output PDF file
   * @param options parameters including font paths and optional passwords
   * @throws PdfGenerateException if an error occurs during PDF generation
   */
  public static void generate(Path excelPath, List<String> sheetNames, Path outputPath,
      PdfGenerateOptions options) throws PdfGenerateException {

    String excelPassword = options.getExcelPassword();
    String pdfPassword = options.getPdfPassword();

    File excelFile = excelPath.toFile();

    try (Workbook workbook = openWorkbook(excelFile, excelPassword);
        PDDocument document = new PDDocument()) {

      Font defaultFont = workbook.getFontAt(0);
      String defaultFontName = defaultFont.getFontName();
      float fontSizePt = defaultFont.getFontHeightInPoints();

      // Step 1: Compute MDW from the theme's Latin minor font.
      // Excel measures column widths using digit widths (0–9) in the Latin body font
      // (e.g. Calibri), NOT the CJK body font stored in font[0] (e.g. 游ゴシック).
      // The theme's minorFont/latin element names the correct font (typically "Calibri").
      int mdw = 0;
      String mdwFontName = defaultFontName; // fallback when no theme is present
      if (workbook instanceof XSSFWorkbook xssfWb) {
        var theme = xssfWb.getStylesSource().getTheme();
        if (theme != null) {
          // getCTTheme() is not public in this POI version; read the theme XML directly.
          try (var themeIs = theme.getPackagePart().getInputStream()) {
            String xml =
                new String(themeIs.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
            // The minorFont's first non-empty typeface is the Latin body font (Calibri etc).
            // Structure: <a:minorFont><a:latin typeface="Calibri" .../><a:ea typeface=""/>...
            var m = java.util.regex.Pattern
                .compile("(?s)<[^>:]+:minorFont[^>]*>.*?typeface=\"([^\"]+)\"").matcher(xml);
            if (m.find()) {
              String tf = m.group(1);
              if (!tf.isBlank()) {
                mdwFontName = tf;
              }
            }
          } catch (Exception ignored) { // NOPMD
            // Keep mdwFontName = defaultFontName
          }
        }
      }
      // Step 2: Determine rendering font (fontManager) and finalise MDW.
      //
      // MDW is always computed at 96 DPI (OOXML standard print resolution).
      // OOXML column widths (§18.3.1.13) are defined in terms of the MDW at 96 DPI,
      // and Excel uses the same 96 DPI basis when computing fit-to-page scales for
      // PDF export — regardless of the screen's physical or logical DPI.
      FontManager fontManager;
      if (options.isUseSystemFonts()) {
        var mdwFontFile = SystemFontLocator.findFontFile(mdwFontName);
        if (mdwFontFile.isPresent()) {
          mdw = SystemFontLocator.computeExcelMdw(mdwFontFile.get(), mdwFontName, fontSizePt);
        }
        var systemFontFile = SystemFontLocator.findFontFile(defaultFontName);
        if (systemFontFile.isEmpty()) {
          // No system font found: try the explicitly specified font as fallback.
          if (options.getRegularFontPath() != null) {
            Path reg = options.getRegularFontPath();
            Path bold = options.getBoldFontPath() != null ? options.getBoldFontPath() : reg;
            fontManager = new FontManager(document, reg, bold);
            if (mdw == 0) {
              mdw = SystemFontLocator.computeExcelMdw(reg, "", fontSizePt);
            }
          } else {
            throw new PdfGenerateException("System font '" + defaultFontName + "' not found. "
                + "Install the font or set regularFontPath as a fallback in "
                + "PdfGenerateOptions.");
          }
        } else {
          Path fontFile = systemFontFile.get();
          if (mdw == 0) {
            mdw = SystemFontLocator.computeExcelMdw(fontFile, defaultFontName, fontSizePt);
          }
          var regularTtf = SystemFontLocator.loadTrueTypeFont(fontFile, defaultFontName);
          if (regularTtf == null) {
            throw new PdfGenerateException(
                "Failed to load font '" + defaultFontName + "' from " + fontFile);
          }
          var boldFontFile = SystemFontLocator.findFontFile(defaultFontName + " Bold");
          var boldTtf = boldFontFile.isPresent()
              ? SystemFontLocator.loadTrueTypeFont(boldFontFile.get(), defaultFontName + " Bold")
              : null;
          // Pass regularFontPath as fallback so that characters not in the system font
          // (e.g. CJK characters in a Calibri workbook) are rendered with the fallback font.
          fontManager = new FontManager(document, regularTtf, boldTtf, options.getRegularFontPath(),
              options.getBoldFontPath());
        }
      } else {
        // Explicit font mode: always use the supplied font at 96 DPI for MDW.
        // 96 DPI is the OOXML standard print resolution and gives consistent results
        // regardless of the screen being used.
        Path reg = java.util.Objects.requireNonNull(options.getRegularFontPath());
        Path bold = options.getBoldFontPath() != null ? options.getBoldFontPath() : reg;
        fontManager = new FontManager(document, reg, bold);
        mdw = SystemFontLocator.computeMdw(reg, "", fontSizePt); // fixed 96 DPI
      }

      Locale dateLocale = options.getDateLocale() != null ? options.getDateLocale()
          : LocaleUtil.getFallbackLocale();
      SheetRenderer renderer = new SheetRenderer(document, fontManager, excelPath, dateLocale, mdw);

      for (String sheetName : sheetNames) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
          throw new PdfGenerateException("Sheet not found: '" + sheetName + "'");
        }
        renderer.render(workbook, sheetIndex);
      }

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
