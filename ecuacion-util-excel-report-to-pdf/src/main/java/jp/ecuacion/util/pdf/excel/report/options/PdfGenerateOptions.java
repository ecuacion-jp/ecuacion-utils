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
import java.util.Locale;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/**
 * Holds optional parameters for PDF generation.
 *
 * <p>Use {@link #builder()} to construct an instance.</p>
 *
 * <p>When {@link Builder#useSystemFonts(boolean) useSystemFonts} is {@code false} (default),
 * {@link Builder#regularFontPath(Path)} is required.
 * When {@code useSystemFonts} is {@code true}, the font matching the workbook's default font
 * is looked up from the OS font directories (including fonts installed by Microsoft Office).
 * If no matching system font is found a {@link
 * jp.ecuacion.util.pdf.excel.report.exception.PdfGenerateException} is thrown at generation time.
 * In that case, specify {@link Builder#regularFontPath(Path)} explicitly as a fallback.</p>
 *
 * <p><strong>Font licensing notice:</strong> when {@code useSystemFonts(true)} is set, the
 * system font is embedded in the output PDF. Ensure that the font's licence permits embedding
 * and distribution before enabling this option.</p>
 */
public class PdfGenerateOptions {

  @Nullable
  private final Path regularFontPath;

  @Nullable
  private final Path boldFontPath;

  private final boolean useSystemFonts;

  @Nullable
  private final String excelPassword;

  @Nullable
  private final String pdfPassword;

  /**
   * The locale used to resolve locale-sensitive built-in date formats (e.g., format ID 14).
   * When {@code null}, {@link Locale#getDefault()} is used at render time.
   */
  @Nullable
  private final Locale dateLocale;

  private PdfGenerateOptions(Builder builder) {
    this.useSystemFonts = builder.useSystemFonts;
    this.regularFontPath = builder.regularFontPath;
    this.boldFontPath = builder.boldFontPath;
    this.excelPassword = builder.excelPassword;
    this.pdfPassword = builder.pdfPassword;
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
   * Returns the path to the regular-weight font file, or {@code null} when
   * {@link #isUseSystemFonts()} is {@code true} and no explicit path was set.
   *
   * @return regular font path, or {@code null}
   */
  @Nullable
  public Path getRegularFontPath() {
    return regularFontPath;
  }

  /**
   * Returns the path to the bold-weight font file, or {@code null} if not set.
   * When {@code null}, the regular font is used for bold text.
   *
   * @return bold font path, or {@code null}
   */
  @Nullable
  public Path getBoldFontPath() {
    return boldFontPath;
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
   * Returns a new {@link Builder} instance.
   *
   * @return builder
   */
  public static Builder builder() {
    return new Builder();
  }

  /**
   * Builder for {@link PdfGenerateOptions}.
   */
  public static class Builder {

    private boolean useSystemFonts = false;

    @Nullable
    private Path regularFontPath;

    @Nullable
    private Path boldFontPath;

    @Nullable
    private String excelPassword;

    @Nullable
    private String pdfPassword;

    @Nullable
    private Locale dateLocale;

    private Builder() {}

    /**
     * Enables system font lookup.
     *
     * <p>When set to {@code true}, the OS font directories (including fonts installed by
     * Microsoft Office) are searched for the font matching the workbook's default font.
     * The found font is used for both PDF rendering (embedded in the output PDF) and
     * accurate column-width calculation (MDW).
     * If no matching font is found, a {@code PdfGenerateException} is thrown at generation time
     * unless {@link #regularFontPath(Path)} is also set as a fallback.</p>
     *
     * <p><strong>Font licensing notice:</strong> the located system font is embedded in the
     * output PDF. Confirm that the font's licence permits embedding and distribution before
     * enabling this option.</p>
     *
     * @param useSystemFonts {@code true} to enable system font lookup (default {@code false})
     * @return this builder
     */
    public Builder useSystemFonts(boolean useSystemFonts) {
      this.useSystemFonts = useSystemFonts;
      return this;
    }

    /**
     * Sets the path to the regular-weight font file.
     *
     * <p>Required when {@link #useSystemFonts(boolean)} is {@code false} (default).
     * Optional when {@code useSystemFonts} is {@code true}; if set it serves as a fallback
     * when the system font cannot be found.</p>
     *
     * @param regularFontPath path to the TTF file used for regular text
     * @return this builder
     */
    public Builder regularFontPath(Path regularFontPath) {
      this.regularFontPath = regularFontPath;
      return this;
    }

    /**
     * Sets the path to the bold-weight font file (optional).
     * When not set, the regular font is used for bold text.
     *
     * @param boldFontPath path to the TTF file used for bold text
     * @return this builder
     */
    public Builder boldFontPath(@Nullable Path boldFontPath) {
      this.boldFontPath = boldFontPath;
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
     * @throws IllegalStateException if {@code useSystemFonts} is {@code false} and
     *     {@code regularFontPath} has not been set
     */
    public PdfGenerateOptions build() {
      if (!useSystemFonts && regularFontPath == null) {
        throw new IllegalStateException(
            "regularFontPath is required when useSystemFonts is false");
      }
      return new PdfGenerateOptions(this);
    }
  }
}
