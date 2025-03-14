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
package jp.ecuacion.util.pdfbox;

import java.io.IOException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.pdfbox.excel.ExcelToPdf;
import org.apache.poi.EncryptedDocumentException;

public class Main {
  public static void main(String[] args) throws Exception {
    new Main().internalMain();
  }

  private void internalMain() throws EncryptedDocumentException, IOException, BizLogicAppException {
    new ExcelToPdf().execute("helloworld.pdf", "local-test/test01.xlsx", new String[] {"Sheet1"});
  }
}
