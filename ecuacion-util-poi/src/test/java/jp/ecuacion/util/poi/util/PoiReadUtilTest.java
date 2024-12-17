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
package jp.ecuacion.util.poi.util;

import static org.junit.jupiter.api.Assertions.assertEquals;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class PoiReadUtilTest {

  private PoiReadUtil reader;

  @BeforeAll
  public static void beforeClass() {}


  @BeforeEach
  public void before() {
    reader = new PoiReadUtil();
  }

  @Test
  public void readFromCellTest()
      throws URISyntaxException, EncryptedDocumentException, IOException {
    String excelPath = new File("src/test/resources").getAbsolutePath() + "/readFromCellTest.xlsx";
    Workbook excel = WorkbookFactory.create(new File(excelPath.toString()), null, true);
    Sheet sheet = excel.getSheet("StringReader");
    int dataCol = 3;
    System.out.println();

    // 表示形式：標準
    int i = 1;
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), null);
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), null);
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123.45");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "あいう");

    // 表示形式：数値
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), null);
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123.45");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "あいう");

    // 表示形式：文字列
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), null);
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "あいう");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "123");

    // 表示形式：日付
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), null);
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "1900/5/2");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "あいう");
    assertEquals(reader.getStringFromCell(sheet.getRow(i++).getCell(dataCol)), "2024/1/1");
  }
}
