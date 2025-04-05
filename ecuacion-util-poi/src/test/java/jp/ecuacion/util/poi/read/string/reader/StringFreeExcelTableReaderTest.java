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
package jp.ecuacion.util.poi.read.string.reader;

import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.reader.concrete.StringFreeExcelTableReader;
import org.apache.poi.EncryptedDocumentException;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class StringFreeExcelTableReaderTest {

  @Test
  public void tableRowSizeNullTest() throws EncryptedDocumentException, AppException, IOException {

    String excelPath = "src/test/resources/StringFreeTableReaderTest.xlsx";
    List<List<String>> rowList = new Reader(6, NoDataString.NULL).read(excelPath);

    Assertions.assertNotNull(rowList);
    Assertions.assertEquals(6, rowList.size());

    // 1行目
    Assertions.assertEquals("header1", rowList.get(0).get(0));
    Assertions.assertEquals("header2", rowList.get(0).get(1));
    Assertions.assertNull(rowList.get(0).get(2));

    // 2行目
    Assertions.assertEquals("data1-1", rowList.get(1).get(0));
    Assertions.assertEquals("data1-2", rowList.get(1).get(1));
    Assertions.assertEquals("data1-3", rowList.get(1).get(2));

    // 3行目
    Assertions.assertEquals("data2-1", rowList.get(2).get(0));
    Assertions.assertNull(rowList.get(2).get(1));
    Assertions.assertEquals("data2-3", rowList.get(2).get(2));

    // 4行目
    Assertions.assertEquals(0, rowList.get(3).size());

    // 5行目
    Assertions.assertEquals("data4-1", rowList.get(4).get(0));
    Assertions.assertEquals("data4-2", rowList.get(4).get(1));
    Assertions.assertEquals("data4-3", rowList.get(4).get(2));

    // 6行目
    Assertions.assertEquals(0, rowList.get(5).size());
  }

  private static class Reader extends StringFreeExcelTableReader {

    public Reader(Integer tableRowSize, @RequireNonnull NoDataString noDataString) {
      super("Sheet1", 3, 2, tableRowSize, 4, ObjectsUtil.paramRequireNonNull(noDataString));
    }
  }
}
