package jp.ecuacion.util.poi.util;

import java.io.File;
import java.io.IOException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
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

    // changesNumberString == true, dataFormat is "text", changesCellsWithDataFormatIsString == false
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
}
