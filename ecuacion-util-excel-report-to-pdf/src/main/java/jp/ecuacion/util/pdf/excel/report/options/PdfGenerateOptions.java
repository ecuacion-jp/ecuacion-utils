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
package jp.ecuacion.util.pdf.excel.report.options;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.jspecify.annotations.Nullable;

/**
 * Holds optional parameters for PDF generation.
 *
 * <p>Use {@link #builderForSystemFonts()} or {@link #builderForExplicitFont(Path)} to construct
 * an instance, depending on how the rendering font should be resolved:</p>
 * <ul>
 *   <li>{@link #builderForSystemFonts()} — the font matching the workbook's default font is
 *       looked up from the OS font directories (including fonts installed by Microsoft Office).
 *       If no matching system font is found a {@link
 *       jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException} is thrown at
 *       generation time, unless {@link Builder#addRegularFontPath(Path)} is also set as a
 *       fallback.</li>
 *   <li>{@link #builderForExplicitFont(Path)} — the given font file is always used, without any
 *       system font lookup.</li>
 * </ul>
 *
 * <p>Characters that cannot be encoded by the resolved font (system-resolved or explicit) are
 * tried against the fonts registered via {@link Builder#addRegularFontPath(Path)}, in
 * registration order (in explicit-font mode, the font passed to {@link
 * #builderForExplicitFont(Path)} is always the first entry). Bold text is additionally tried
 * against fonts registered via {@link Builder#addBoldFontPath(Path)}, in registration order,
 * before falling through to the regular fonts above.</p>
 *
 * <p><strong>Font licensing notice:</strong> when {@link #builderForSystemFonts()} is used, the
 * system font is embedded in the output PDF. Ensure that the font's licence permits embedding
 * and distribution before enabling this option.</p>
 */
public class PdfGenerateOptions {

  private final List<Path> regularFontPaths;

  private final List<Path> boldFontPaths;

  private final boolean useSystemFonts;

  @Nullable
  private final String excelPassword;

  @Nullable
  private final String pdfPassword;

  @Nullable
  private final String pdfOwnerPassword;

  /**
   * The locale used to resolve locale-sensitive built-in date formats (e.g., format ID 14).
   * When {@code null}, {@link Locale#getDefault()} is used at render time.
   */
  @Nullable
  private final Locale dateLocale;

  private PdfGenerateOptions(Builder builder) {
    this.useSystemFonts = builder.useSystemFonts;
    this.regularFontPaths = List.copyOf(builder.regularFontPaths);
    this.boldFontPaths = List.copyOf(builder.boldFontPaths);
    this.excelPassword = builder.excelPassword;
    this.pdfPassword = builder.pdfPassword;
    this.pdfOwnerPassword = builder.pdfOwnerPassword;
    this.dateLocale = builder.dateLocale;
  }

  /**
   * Returns {@code true} if system font lookup is enabled.
   *
   * @return whether system fonts are used
   */
  public boolean isUseSystemFonts() {
    return useSystemFonts;
  }

  /**
   * Returns the regular-weight font files registered via {@link Builder#addRegularFontPath(Path)},
   * in registration order. In explicit-font mode, the first entry is the font passed to
   * {@link #builderForExplicitFont(Path)}. Possibly empty in system-fonts mode.
   *
   * @return regular font paths, possibly empty, never {@code null}
   */
  public List<Path> getRegularFontPaths() {
    return regularFontPaths;
  }

  /**
   * Returns the bold-weight font files registered via {@link Builder#addBoldFontPath(Path)}, in
   * registration order. When empty, {@link #getRegularFontPaths()} is used for bold text as well.
   *
   * @return bold font paths, possibly empty, never {@code null}
   */
  public List<Path> getBoldFontPaths() {
    return boldFontPaths;
  }

  /**
   * Returns the password used to open the Excel file, or {@code null} if none.
   *
   * @return Excel file password
   */
  @Nullable
  public String getExcelPassword() {
    return excelPassword;
  }

  /**
   * Returns the password to protect the output PDF, or {@code null} if none.
   *
   * @return PDF password
   */
  @Nullable
  public String getPdfPassword() {
    return pdfPassword;
  }

  /**
   * Returns the owner password for the output PDF, or {@code null} if none was explicitly set.
   *
   * <p>When {@code null}, the value of {@link #getPdfPassword()} is used as the owner password.</p>
   *
   * @return PDF owner password, or {@code null}
   */
  @Nullable
  public String getPdfOwnerPassword() {
    return pdfOwnerPassword;
  }

  /**
   * Returns the locale used to resolve locale-sensitive built-in date formats,
   * or {@code null} if the JVM default locale ({@link Locale#getDefault()}) should be used.
   *
   * @return date locale, or {@code null}
   */
  @Nullable
  public Locale getDateLocale() {
    return dateLocale;
  }

  /**
   * Returns a new {@link Builder} configured to resolve the rendering font from the OS font
   * directories (see {@link #isUseSystemFonts()}). Fonts registered via {@link
   * Builder#addRegularFontPath(Path)} stay optional, and act as fallback fonts, tried in
   * registration order, when the system font cannot be found or lacks a glyph.
   *
   * @return builder
   */
  public static Builder builderForSystemFonts() {
    Builder b = new Builder();
    b.useSystemFonts = true;
    return b;
  }

  /**
   * Returns a new {@link Builder} configured to always render with the given font file,
   * without any system font lookup.
   *
   * @param regularFontPath path to the TTF/TTC file used for regular text
   * @return builder
   */
  public static Builder builderForExplicitFont(Path regularFontPath) {
    Builder b = new Builder();
    b.useSystemFonts = false;
    b.addRegularFontPath(regularFontPath);
    return b;
  }

  /**
   * Builder for {@link PdfGenerateOptions}.
   */
  public static class Builder {

    private boolean useSystemFonts = false;

    private final List<Path> regularFontPaths = new ArrayList<>();

    private final List<Path> boldFontPaths = new ArrayList<>();

    @Nullable
    private String excelPassword;

    @Nullable
    private String pdfPassword;

    @Nullable
    private String pdfOwnerPassword;

    @Nullable
    private Locale dateLocale;

    private Builder() {}

    /**
     * Registers a regular-weight font file, tried (in the order registered) when a character
     * cannot be encoded by an earlier-registered regular font. Can be called multiple times to
     * cover several scripts/languages not present in an earlier font (e.g. a workbook mixing
     * Japanese, Korean, and Arabic text where no single font covers all three).
     *
     * <p>In explicit-font mode, the font passed to {@link
     * PdfGenerateOptions#builderForExplicitFont(Path)} is always registered first; calling this
     * afterwards adds further fallback fonts. In system-fonts mode, every font registered here
     * is used as a fallback, tried in registration order, after the system-resolved font.</p>
     *
     * @param regularFontPath path to the TTF/TTC file used for regular text
     * @return this builder
     */
    public Builder addRegularFontPath(Path regularFontPath) {
      this.regularFontPaths.add(regularFontPath);
      return this;
    }

    /**
     * Registers a bold-weight font file, tried (in the order registered) when a character
     * cannot be encoded by an earlier-registered bold font. Can be called multiple times, in
     * the same way as {@link #addRegularFontPath(Path)}.
     *
     * <p>Bold text is tried against these fonts first, then falls through to {@link
     * #addRegularFontPath(Path)}'s fonts (in registration order) for any character not covered.
     * When this is never called, bold text is rendered entirely with the regular fonts.</p>
     *
     * @param boldFontPath path to the TTF/TTC file used for bold text
     * @return this builder
     */
    public Builder addBoldFontPath(Path boldFontPath) {
      this.boldFontPaths.add(boldFontPath);
      return this;
    }

    /**
     * Sets the password used to open the Excel file.
     *
     * @param excelPassword Excel file password
     * @return this builder
     */
    public Builder excelPassword(@Nullable String excelPassword) {
      this.excelPassword = excelPassword;
      return this;
    }

    /**
     * Sets the password to protect the output PDF.
     *
     * @param pdfPassword PDF password
     * @return this builder
     */
    public Builder pdfPassword(@Nullable String pdfPassword) {
      this.pdfPassword = pdfPassword;
      return this;
    }

    /**
     * Sets the owner password for the output PDF.
     *
     * <p>The owner password controls who can modify the PDF's security settings
     * (e.g. add print or copy restrictions using an external tool after generation).
     * When not set, {@code pdfPassword} is used as the owner password as well.</p>
     *
     * <p>Set this when the person who generates the PDF and the person who later
     * configures print/copy restrictions are different — for example, a system that
     * generates the PDF with {@code pdfPassword} for end-user access, while an
     * administrator applies restrictions using a separately managed {@code pdfOwnerPassword}.</p>
     *
     * @param pdfOwnerPassword PDF owner password
     * @return this builder
     */
    public Builder pdfOwnerPassword(@Nullable String pdfOwnerPassword) {
      this.pdfOwnerPassword = pdfOwnerPassword;
      return this;
    }

    /**
     * Sets the locale used to resolve locale-sensitive built-in date formats (e.g., format ID 14).
     * When not set, {@link Locale#getDefault()} is used at render time.
     *
     * @param dateLocale locale for date format resolution
     * @return this builder
     */
    public Builder dateLocale(@Nullable Locale dateLocale) {
      this.dateLocale = dateLocale;
      return this;
    }

    /**
     * Builds a new {@link PdfGenerateOptions} instance.
     *
     * @return options
     */
    public PdfGenerateOptions build() {
      return new PdfGenerateOptions(this);
    }
  }
}
