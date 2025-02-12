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
package jp.ecuacion.util.poi.excel.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable.ContextContainer;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Provides excel writing related {@code apache POI} utility methods.
 */
public class ExcelWriteUtil {

  private DetailLogger detailLog = new DetailLogger(this);

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

  /**
   * Opens the excel file and returns {@code Workbook} object.
   * 
   * @param filePath filePath
   * @return workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public Workbook openForWrite(String filePath) throws EncryptedDocumentException, IOException {
    return WorkbookFactory.create(new FileInputStream(filePath));
  }

  /**
   * Opens the excel file and returns {@code Workbook} object.
   * 
   * @param filePath filePath
   * @return workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public FileOutputStream openForOutput(String filePath)
      throws EncryptedDocumentException, IOException {
    return new FileOutputStream(filePath);
  }

  /**
   * Opens the excel file and returns {@code Workbook} object.
   */
  public void saveToFile(Workbook workbook, FileOutputStream out)
      throws EncryptedDocumentException, IOException {
    workbook.write(out);
  }

  /**
   * Gets ready to write table data.
   */
  public <T> ContextContainer getReadyToWriteTableData(ExcelTableWriter<T> writer,
      Workbook workbook, String sheetName) throws BizLogicAppException {

    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("starting to write excel file.");
    detailLog.debug("sheet name :" + sheetName);

    Sheet sheet = workbook.getSheet(sheetName);

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", sheetName);
    }

    int poiBasisTableStartColumnNumber = writer.getPoiBasisDeterminedTableStartColumnNumber();
    int poiBasisTableStartRowNumber = writer.getPoiBasisDeterminedTableStartRowNumber(sheet);

    // Skip the header line if the writer is OneLineHeaderFormat
    if (this instanceof IfFormatOneLineHeaderExcelTable) {
      poiBasisTableStartRowNumber++;
    }

    return new ContextContainer(sheet, poiBasisTableStartRowNumber, poiBasisTableStartColumnNumber,
        null, null);
  }

  /**
   * Provides common procedure for write one line of a table.
   */
  public <T> void writeTableLine(ExcelTableWriter<T> writer, ContextContainer context,
      int rowNumber, List<T> columnList) {

    if (context.sheet.getRow(rowNumber) == null) {
      context.sheet.createRow(rowNumber);
    }

    Row row = context.sheet.getRow(rowNumber);

    for (int colNumber =
        context.poiBasisTableStartColumnNumber; colNumber < context.poiBasisTableStartColumnNumber
            + columnList.size(); colNumber++) {

      T sourceCellData = columnList.get(colNumber - context.poiBasisTableStartColumnNumber);

      if (row.getCell(colNumber) == null) {
        row.createCell(colNumber);
      }

      Cell destCell = row.getCell(colNumber);

      writer.writeToCell(colNumber - context.poiBasisTableStartColumnNumber, sourceCellData,
          destCell);
    }

  }
}
