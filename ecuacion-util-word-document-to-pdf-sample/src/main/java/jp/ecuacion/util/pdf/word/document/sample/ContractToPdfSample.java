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
package jp.ecuacion.util.pdf.word.document.sample;

import java.io.File;
import java.nio.file.Path;
import java.util.Objects;
import jp.ecuacion.util.pdf.word.document.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.word.document.util.WordToPdfUtil;

/**
 * Converts a contract written in Word ({@code test-data/contract.docx}) to PDF.
 *
 * <p>{@code test-data/contract.docx.pdf} is the PDF file Word for Mac exported from the same
 *     file, for comparison.</p>
 */
public class ContractToPdfSample {
  public static void main(String[] args) throws Exception {
    new File("target/test-result").mkdirs();

    // The fonts the document specifies (Yu Mincho, Century etc.) are looked up from the system.
    // NotoSansJP is used for the text whose fonts are not found.
    var reg = Objects.requireNonNull(
        ContractToPdfSample.class.getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf"));
    var bold = Objects.requireNonNull(
        ContractToPdfSample.class.getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf"));
    PdfGenerateOptions options = PdfGenerateOptions.builderForSystemFonts()
        .addRegularFontPath(Path.of(reg.toURI()))
        .addBoldFontPath(Path.of(bold.toURI()))
        .build();

    WordToPdfUtil.generate(Path.of("test-data/contract.docx"),
        Path.of("target/test-result/contract.pdf"), options);

    // To render with a font file whatever fonts the document specifies:
    PdfGenerateOptions explicit = PdfGenerateOptions.builderForExplicitFont(Path.of(reg.toURI()))
        .addBoldFontPath(Path.of(bold.toURI()))
        .build();

    WordToPdfUtil.generate(Path.of("test-data/contract.docx"),
        Path.of("target/test-result/contract-notosans.pdf"), explicit);
  }
}
