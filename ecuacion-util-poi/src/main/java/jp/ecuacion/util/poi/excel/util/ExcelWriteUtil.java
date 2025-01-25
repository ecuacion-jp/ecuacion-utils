package jp.ecuacion.util.poi.excel.util;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Provides excel writing related {@code apache POI} utility methods.
 */
public class ExcelWriteUtil {
  
  /**
   * Creates new workbook with adding sheet of name {@code sheetName}.
   * 
   * @param sheetName sheetName
   * @return Workbook
   */
  public Workbook createWorkbookWithSheet(String sheetName) {
    Workbook wb = new XSSFWorkbook();
    wb.createSheet(sheetName);
    
    return wb;
  }
}
