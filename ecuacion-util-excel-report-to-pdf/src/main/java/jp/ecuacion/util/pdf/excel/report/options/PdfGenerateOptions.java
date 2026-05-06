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

import org.jspecify.annotations.Nullable;

/**
 * Holds optional parameters for PDF generation.
 *
 * <p>Use {@link #builder()} to construct an instance.</p>
 */
public class PdfGenerateOptions {

  @Nullable
  private final String excelPassword;

  @Nullable
  private final String pdfPassword;

  private PdfGenerateOptions(Builder builder) {
    this.excelPassword = builder.excelPassword;
    this.pdfPassword = builder.pdfPassword;
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

    @Nullable
    private String excelPassword;

    @Nullable
    private String pdfPassword;

    private Builder() {}

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
     * Builds a new {@link PdfGenerateOptions} instance.
     *
     * @return options
     */
    public PdfGenerateOptions build() {
      return new PdfGenerateOptions(this);
    }
  }
}
