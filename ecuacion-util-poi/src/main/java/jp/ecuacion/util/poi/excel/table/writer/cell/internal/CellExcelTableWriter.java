package jp.ecuacion.util.poi.excel.table.writer.cell.internal;

import jp.ecuacion.util.poi.excel.table.writer.core.ExcelTableWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellUtil;

public abstract class CellExcelTableWriter extends ExcelTableWriter<Cell> {

  public CellExcelTableWriter(String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  protected void writeToCell(Cell sourceCellData, Cell destCell) {
    CellCopyPolicy policy = new CellCopyPolicy();
    policy.setCopyCellFormula(false);
    
    CellUtil.copyCell(sourceCellData, destCell, policy, null);
  }
}
