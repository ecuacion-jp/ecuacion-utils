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
import java.time.format.DateTimeFormatter;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class ExcelReadUtilTest {

  private ExcelReadUtil reader;

  @BeforeAll
  public static void beforeClass() {}


  @BeforeEach
  public void before() {
    reader = new ExcelReadUtil();
  }

  @Test
  public void readFromCellTest()
      throws URISyntaxException, EncryptedDocumentException, IOException, BizLogicAppException {
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
    row++; // （テストなし）
    assertEquals("1.23457E11", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1234567890", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12345.12346", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("36548", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++;
    row++; // （テストなし）×2
    assertEquals("0.5242592593", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("36548.52426", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
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
    row++;
    row++; // （テストなし）×2
    assertEquals("1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("36549", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));

    // 表示形式：日付
    row = START_ROW;
    dataCol = 11;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900-05-02", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900-05-02", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; // （テストなし）
    assertEquals("+3002035-06-10", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("+3382032-01-27", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1933-10-18", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000-01-23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000-01-23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000-01-23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12:34:56", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol),
        DateTimeFormatter.ofPattern("HH:mm:ss")));
    assertEquals("2000/1/23 12:34:56", reader.getStringFromCell(
        sheet.getRow(row++).getCell(dataCol), DateTimeFormatter.ofPattern("yyyy/M/dd HH:mm:ss")));
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1900-01-01", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));

    // 表示形式：文字列
    row = START_ROW;
    dataCol = 14;
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals(null, reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("123.45", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    row++; // （テストなし）
    assertEquals("123456789012", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("1234567890.12", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12345.1234567", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000/1/23", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("23-Jan-00", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000年1月", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("12:34:56", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("2000/1/23 12:34:56",
        reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("あいう", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
    assertEquals("=$A$1", reader.getStringFromCell(sheet.getRow(row++).getCell(dataCol)));
  }
}
