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

import java.io.File;
import java.io.IOException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

public class ExcelWriteUtilTest {

  private ExcelWriteUtil writer;

  @BeforeAll
  public static void beforeClass() {}


  @BeforeEach
  public void before() {
    writer = new ExcelWriteUtil();
  }

  @Test
  public void getReadyToEvaluateFormulaTest()
      throws EncryptedDocumentException, IOException, BizLogicAppException {
    String excelPath =
        new File("src/test/resources").getAbsolutePath() + "/ExcelWriteUtilTest.xlsx";

    Sheet sheet = writer.openForWrite(excelPath).getSheet("getReadyToEvaluateFormulaTest");
    Cell cell;

    // numberString

    // changesNumberString == false, dataFormat is "number"
    cell = sheet.getRow(1).getCell(2);
    writer.getReadyToEvaluateFormula(cell, false, false, false, null);
    // unchanged
    Assertions.assertEquals(CellType.STRING, cell.getCellType());
    Assertions.assertEquals("1", cell.getStringCellValue());

    // changesNumberString == true, dataFormat is "number"
    cell = sheet.getRow(1).getCell(2);
    writer.getReadyToEvaluateFormula(cell, true, false, false, null);
    // changed
    Assertions.assertEquals(CellType.NUMERIC, cell.getCellType());
    Assertions.assertEquals(1, cell.getNumericCellValue());

    // changesNumberString == true, dataFormat is "text", changesCellsWithDataFormatIsString ==
    // false
    cell = sheet.getRow(2).getCell(2);
    writer.getReadyToEvaluateFormula(cell, true, false, false, null);
    // unchanged
    Assertions.assertEquals(CellType.STRING, cell.getCellType());
    Assertions.assertEquals("1", cell.getStringCellValue());

    // changesNumberString == true, dataFormat is "text", changesCellsWithDataFormatIsString == true
    cell = sheet.getRow(2).getCell(2);
    writer.getReadyToEvaluateFormula(cell, true, false, true, null);
    // changed
    Assertions.assertEquals(CellType.NUMERIC, cell.getCellType());
    Assertions.assertEquals(1, cell.getNumericCellValue());


    // dateString

    // changesDateString == false, dataFormat is "number"
    cell = sheet.getRow(3).getCell(2);
    writer.getReadyToEvaluateFormula(cell, false, false, false, new String[] {"yyyy/MM/dd"});
    // unchanged
    Assertions.assertEquals(CellType.STRING, cell.getCellType());
    Assertions.assertEquals("2025/01/01", cell.getStringCellValue());

    // changesDateString == true, dataFormat is "number"
    cell = sheet.getRow(3).getCell(2);
    writer.getReadyToEvaluateFormula(cell, false, true, false, new String[] {"yyyy/MM/dd"});
    // changed
    Assertions.assertEquals(CellType.NUMERIC, cell.getCellType());
    Assertions.assertEquals(45658, cell.getNumericCellValue());

    // changesDateString == true, dataFormat is "text", changesCellsWithDataFormatIsString == false
    cell = sheet.getRow(4).getCell(2);
    writer.getReadyToEvaluateFormula(cell, false, true, false, new String[] {"yyyy/MM/dd"});
    // unchanged
    Assertions.assertEquals(CellType.STRING, cell.getCellType());
    Assertions.assertEquals("2025/01/01", cell.getStringCellValue());

    // changesDateString == true, dataFormat is "text", changesCellsWithDataFormatIsString == true
    cell = sheet.getRow(4).getCell(2);
    writer.getReadyToEvaluateFormula(cell, false, true, true, new String[] {"yyyy/MM/dd"});
    // changed
    Assertions.assertEquals(CellType.NUMERIC, cell.getCellType());
    Assertions.assertEquals(45658, cell.getNumericCellValue());


    // CellType != STRING
    cell = sheet.getRow(5).getCell(2);
    writer.getReadyToEvaluateFormula(cell, true, true, true, new String[] {"yyyy/MM/dd"});
    // ignored
    Assertions.assertEquals(CellType.NUMERIC, cell.getCellType());
    Assertions.assertEquals(1, cell.getNumericCellValue());
  }

  @Test
  public void evaluateFormulaTest()
      throws EncryptedDocumentException, IOException, BizLogicAppException {
    String excelPath =
        new File("src/test/resources").getAbsolutePath() + "/ExcelWriteUtilTest.xlsx";
    Workbook wb = writer.openForWrite(excelPath);
    Sheet sheet = wb.getSheet("evaluateFormulaTest");

    // #NAME?
    try {
      writer.evaluateFormula(sheet.getRow(3).getCell(1), "testfile");
      Assertions.fail();
    } catch (BizLogicAppException ex) {
      Assertions.assertEquals("jp.ecuacion.util.poi.excel.ExcelWriteUtil.DetailUnknown.message",
          ex.getMessageId());
    }

  }
}
