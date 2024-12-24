package jp.ecuacion.util.poi.write.cell.writer;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.write.core.reader.TableWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;

/**
 * 
 */
public class CellFixedTableWriter extends TableWriter {
  private String[] headerLabels;

  /**
   * 
   * @param sheetName
   * @param headerLabels
   * @param tableStartRowNumber
   * @param tableStartColumnNumber
   */
  public CellFixedTableWriter(@RequireNonnull String sheetName, String[] headerLabels,
      Integer tableStartRowNumber, int tableStartColumnNumber) {

    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.headerLabels = headerLabels;
  }

  /**
   * 
   * @param destFilePath 
   * @param dataList
   */
  public void write(String templateFilepath, String destFilePath, List<List<Cell>> dataList) throws Exception {

    Workbook excel = WorkbookFactory.create(new File(templateFilepath), null, false);
    Sheet sheet = excel.getSheet(getSheetName());

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", templateFilepath, getSheetName());
    }

    int rowNum = 0;
    int colNum = 0;

    for (rowNum = getPoiBasisTableStartRowNumber() + 1; rowNum < getPoiBasisTableStartRowNumber()
        + 1 + dataList.size(); rowNum++) {

      List<Cell> list = dataList.get(rowNum - (getPoiBasisTableStartRowNumber() + 1));

      if (sheet.getRow(rowNum) == null) {
        sheet.createRow(rowNum);
      }

      Row row = sheet.getRow(rowNum);

      for (colNum =
          getPoiBasisTableStartColumnNumber(); colNum < getPoiBasisTableStartColumnNumber()
              + headerLabels.length; colNum++) {
        
        Cell sourceCell = list.get(colNum - getPoiBasisTableStartColumnNumber());
        
        if (row.getCell(colNum) == null) {
          row.createCell(colNum);
        }
        
        Cell destCell = row.getCell(colNum);

        CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyCellFormula(false);
        
        CellUtil.copyCell(sourceCell, destCell, policy, null);
      }
    }
    
    // 出力用のストリームを用意
    FileOutputStream out = new FileOutputStream(destFilePath);
 
    // ファイルへ出力
    excel.write(out);
  }
}
