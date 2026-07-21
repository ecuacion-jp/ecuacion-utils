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
 *       generation time, unless {@link Builder#regularFontPath(Path)} is also set as a
 *       fallback.</li>
 *   <li>{@link #builderForExplicitFont(Path)} — the given font file is always used, without any
 *       system font lookup.</li>
 * </ul>
 *
 * <p>Characters that cannot be encoded by the resolved font (system-resolved or explicit) are
 * tried against fallback fonts in order: first {@link Builder#regularFontPath(Path)}/
 * {@link Builder#boldFontPath(Path)} (when set — always set in explicit-font mode), then any
 * fonts registered via {@link Builder#addFallbackFont(Path, Path)}, which can be called
 * multiple times to cover several scripts/languages not present in the primary font.</p>
 *
 * <p><strong>Font licensing notice:</strong> when {@link #builderForSystemFonts()} is used, the
 * system font is embedded in the output PDF. Ensure that the font's licence permits embedding
 * and distribution before enabling this option.</p>
 */
public class PdfGenerateOptions {

  @Nullable
  private final Path regularFontPath;

  @Nullable
  private final Path boldFontPath;

  private final List<FallbackFont> additionalFallbackFonts;

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
    this.regularFontPath = builder.regularFontPath;
    this.boldFontPath = builder.boldFontPath;
    this.additionalFallbackFonts = List.copyOf(builder.additionalFallbackFonts);
    this.excelPassword = builder.excelPassword;
    this.pdfPassword = builder.pdfPassword;
    this.pdfOwnerPassword = builder.pdfOwnerPassword;
    this.dateLocale = builder.dateLocale;
  }

  /**
   * A regular/bold font path pair registered as an additional fallback font via
   * {@link Builder#addFallbackFont(Path, Path)}.
   *
   * @param regularFontPath path to the TTF/TTC file used for regular text
   * @param boldFontPath    path to the TTF/TTC file used for bold text, or {@code null} to
   *                        fall back to {@code regularFontPath}
   */
  public record FallbackFont(Path regularFontPath, @Nullable Path boldFontPath) {
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
   * Returns the path to the regular-weight font file. Always non-null when this instance was
   * built via {@link #builderForExplicitFont(Path)}; may be {@code null} when built via
   * {@link #builderForSystemFonts()} and no fallback path was set.
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
   * Returns the additional fallback fonts registered via
   * {@link Builder#addFallbackFont(Path, Path)}, in registration order.
   * Tried, in order, after {@link #getRegularFontPath()}/{@link #getBoldFontPath()} when a
   * character cannot be encoded by the primary font.
   *
   * @return additional fallback fonts, possibly empty, never {@code null}
   */
  public List<FallbackFont> getAdditionalFallbackFonts() {
    return additionalFallbackFonts;
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
   * directories (see {@link #isUseSystemFonts()}). {@link Builder#regularFontPath(Path)} stays
   * optional, and acts as the first fallback font when set.
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
    b.regularFontPath = regularFontPath;
    return b;
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

    private final List<FallbackFont> additionalFallbackFonts = new ArrayList<>();

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
     * Sets the path to the regular-weight font file.
     *
     * <p>Already set (and required) when this builder was created via
     * {@link PdfGenerateOptions#builderForExplicitFont(Path)}. When the builder was created via
     * {@link PdfGenerateOptions#builderForSystemFonts()}, calling this is optional and sets the
     * first fallback font, used when the system font cannot be found or lacks a glyph.</p>
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
     * Registers an additional fallback font, tried (in the order registered) after
     * {@link #regularFontPath(Path)}/{@link #boldFontPath(Path)} when a character cannot be
     * encoded by the primary font. Can be called multiple times to cover several
     * scripts/languages not present in the primary font (e.g. a workbook mixing Japanese,
     * Korean, and Arabic text where no single font covers all three).
     *
     * @param regularFontPath path to the TTF/TTC file used for regular text
     * @param boldFontPath    path to the TTF/TTC file used for bold text, or {@code null} to
     *                        fall back to {@code regularFontPath}
     * @return this builder
     */
    public Builder addFallbackFont(Path regularFontPath, @Nullable Path boldFontPath) {
      this.additionalFallbackFonts.add(new FallbackFont(regularFontPath, boldFontPath));
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
