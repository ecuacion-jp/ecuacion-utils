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
package jp.ecuacion.util.pdf.excel.report.sample;

import java.io.File;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.pdf.excel.report.options.PdfGenerateOptions;
import jp.ecuacion.util.pdf.excel.report.util.ExcelToPdfUtil;

public class Test {
  @SuppressWarnings("null")
  public static void main(String[] args) throws Exception {
    new File("target/test-result").mkdirs();

    var reg = Test.class.getResource("/fonts/NotoSansJP/NotoSansJP-Regular.ttf");
    var bold = Test.class.getResource("/fonts/NotoSansJP/NotoSansJP-Bold.ttf");
    PdfGenerateOptions options = PdfGenerateOptions.builder()
        .useSystemFonts(true)
        .regularFontPath(Path.of(reg.toURI()))  // system font が見つからない場合のフォールバック
        .boldFontPath(Path.of(bold.toURI()))
        .build();

    ExcelToPdfUtil.generate(Path.of("test-data/invoice-1.xlsx"), List.of("invoice"),
        Path.of("target/test-result/invoice-1.pdf"), options);

    ExcelToPdfUtil.generate(Path.of("test-data/invoice-2.xlsx"), List.of("invoice"),
        Path.of("target/test-result/invoice-2.pdf"), options);

    ExcelToPdfUtil.generate(Path.of("test-data/invoice-3.xlsx"), List.of("invoice"),
        Path.of("target/test-result/invoice-3.pdf"), options);
  }
}
