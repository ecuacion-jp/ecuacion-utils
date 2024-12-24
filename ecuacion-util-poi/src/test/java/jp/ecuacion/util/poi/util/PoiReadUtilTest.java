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
    System.out.println();
    final int START_ROW = 2;

    // 表示形式：標準
    int row = START_ROW;
    int dataCol = 5;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123.45", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; //（テストなし）
    assertEquals("1.23457E11", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1234567890", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12345.12346",reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("36548", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; row++; //（テストなし）×2
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));

    // 表示形式：数値
    row = START_ROW;
    dataCol = 8;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123.5", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123456789012", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1234567890.12", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12345.1234567", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("36548", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; row++; //（テストなし）×2
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));

    // 表示形式：日付
    row = START_ROW;
    dataCol = 11;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900/5/2", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900/5/2", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; //（テストなし）
    assertEquals("5881510/8/3", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("3382032/1/27", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1933/10/18", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000/1/23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("[$-409]23\\-000\\-00;@", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000\"年\"0\"月\";@", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900/1/1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));

    // 表示形式：文字列
    row = START_ROW;
    dataCol = 14;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123.45", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; //（テストなし）
    assertEquals("123456789012", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1234567890.12", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12345.1234567", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000/1/23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("23-Jan-00", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000年1月", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("=$A$1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
  }
}
