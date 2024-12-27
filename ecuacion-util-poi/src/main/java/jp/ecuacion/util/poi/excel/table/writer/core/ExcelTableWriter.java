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
package jp.ecuacion.util.poi.excel.table.writer.core;

import jakarta.annotation.Nullable;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.IfOneLineHeaderFormatExcelTable;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Is a parent of excel table writer classes.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTableWriter<T> extends ExcelTable<T> implements IfExcelTable<T> {

  private DetailLogger detailLog = new DetailLogger(this);

  /**
   * Constructs a new instance with the sheet name, the position of the excel table.
   * 
   * <p>About the params {@code sheetName}, {@code tableStartRowNumber} and
   *     {@code tableStartColumnNumber},
   *     see {@link ExcelTable#ExcelTable(String, Integer, int)}.</p>
   */
  public ExcelTableWriter(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  /**
   * write table data to the designated excel file.
   * 
   * @param destFilePath destFilePath
   * @param dataList dataList
   */
  public void write(String templateFilePath, String destFilePath, List<List<T>> dataList)
      throws Exception {

    List<List<String>> headerList = getHeaderList(templateFilePath, dataList.get(0).size());

    validateHeader(headerList);

    writeTableValues(templateFilePath, destFilePath, dataList);
  }

  /**
   * Obtains header list.
   * 
   * @param tableColumnSize tableColumnSize
   * @return the list
   * @throws IOException IOException
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  protected abstract List<List<String>> getHeaderList(String templateFilePath, int tableColumnSize)
      throws EncryptedDocumentException, AppException, IOException;

  private void writeTableValues(String templateFilePath, String destFilePath,
      List<List<T>> dataList) throws FileNotFoundException, IOException, BizLogicAppException {

    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("starting to write excel file.");
    detailLog.debug(
        "template file name  :" + ((templateFilePath == null) ? "(none)" : templateFilePath));
    detailLog.debug("sheet name :" + getSheetName());

    Workbook excel = (templateFilePath == null) ? new XSSFWorkbook()
        : WorkbookFactory.create(new File(templateFilePath), null, false);
    
    if (templateFilePath == null) {
      excel.createSheet(getSheetName());
    }
    
    Sheet sheet = excel.getSheet(getSheetName());

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", templateFilePath, getSheetName());
    }

    int poiBasisTableStartColumnNumber = getPoiBasisDeterminedTableStartColumnNumber();
    int poiBasisTableStartRowNumber = getPoiBasisDeterminedTableStartRowNumber(sheet);

    // Skip the header line if the writer is OneLineHeaderFormat
    if (this instanceof IfOneLineHeaderFormatExcelTable) {
      poiBasisTableStartRowNumber++;
    }

    for (int rowNum = poiBasisTableStartRowNumber; rowNum < poiBasisTableStartRowNumber
        + dataList.size(); rowNum++) {

      List<T> list = dataList.get(rowNum - poiBasisTableStartRowNumber);

      if (sheet.getRow(rowNum) == null) {
        sheet.createRow(rowNum);
      }

      Row row = sheet.getRow(rowNum);

      for (int colNum = poiBasisTableStartColumnNumber; colNum < poiBasisTableStartColumnNumber
          + dataList.get(0).size(); colNum++) {

        T sourceCellData = list.get(colNum - poiBasisTableStartColumnNumber);

        if (row.getCell(colNum) == null) {
          row.createCell(colNum);
        }

        Cell destCell = row.getCell(colNum);

        writeToCell(sourceCellData, destCell);
      }
    }

    // 出力用のストリームを用意。ファイルへ出力
    try (FileOutputStream out = new FileOutputStream(destFilePath);) {
      excel.write(out);
    }
  }

  /**
   * writes cell data to the cell.
   * 
   * @param sourceCellData sourceCellData
   * @param destCell destCell
   */
  protected abstract void writeToCell(T sourceCellData, Cell destCell);
}
