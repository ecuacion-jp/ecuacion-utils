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
package jp.ecuacion.util.poi.excel.table.reader.concrete;

import java.util.List;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class CellOneLineHeaderExcelTableReaderTest {

  private String filename = this.getClass().getSimpleName() + ".xlsx";

  @Test
  public void normalTableTest() throws Exception {
    String origExcelPath = "src/test/resources/" + filename;
    final String[] HEADER_LABELS = new String[] {"header1", "header2", "header3"};

    List<List<Cell>> rowList =
        new CellOneLineHeaderExcelTableReader("copy-from", HEADER_LABELS, 2, 1, null)
            .read(origExcelPath);

    Assertions.assertEquals("data1-1", ExcelReadUtil.getStringFromCell(rowList.get(0).get(0)));
    Assertions.assertEquals("data1-2", ExcelReadUtil.getStringFromCell(rowList.get(0).get(1)));
    Assertions.assertEquals("data1-3", ExcelReadUtil.getStringFromCell(rowList.get(0).get(2)));

    Assertions.assertEquals("data2-1", ExcelReadUtil.getStringFromCell(rowList.get(1).get(0)));
    Assertions.assertEquals("data2-2", ExcelReadUtil.getStringFromCell(rowList.get(1).get(1)));
    Assertions.assertEquals("data2-3", ExcelReadUtil.getStringFromCell(rowList.get(1).get(2)));
  }

  @Test
  public void verticalHeaderTableTest() throws Exception {
    String origExcelPath = "src/test/resources/" + filename;
    final String[] HEADER_LABELS = new String[] {"header1", "header2", "header3"};

    List<List<Cell>> rowList =
        new CellOneLineHeaderExcelTableReader("copy-from", HEADER_LABELS, null, 17, null)
            .isVerticalAndHorizontalOpposite(true).read(origExcelPath);

    Assertions.assertEquals("data1-1", ExcelReadUtil.getStringFromCell(rowList.get(0).get(0)));
    Assertions.assertEquals("data1-2", ExcelReadUtil.getStringFromCell(rowList.get(0).get(1)));
    Assertions.assertEquals("data1-3", ExcelReadUtil.getStringFromCell(rowList.get(0).get(2)));

    Assertions.assertEquals("data2-1", ExcelReadUtil.getStringFromCell(rowList.get(1).get(0)));
    Assertions.assertEquals("data2-2", ExcelReadUtil.getStringFromCell(rowList.get(1).get(1)));
    Assertions.assertEquals("data2-3", ExcelReadUtil.getStringFromCell(rowList.get(1).get(2)));
  }
}
