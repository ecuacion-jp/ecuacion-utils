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
package jp.ecuacion.util.poi.excel.table.writer;

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
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;
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
public abstract class ExcelTableWriter<T> extends ExcelTable<T> implements IfExcelTableWriter<T> {

  private DetailLogger detailLog = new DetailLogger(this);

  /**
   * Constructs a new instance with the sheet name, the position of the excel table.
   * 
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public ExcelTableWriter(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  /**
   * Writes table data to the specified excel file.
   * 
   * <p>{@code data} is written to {@code destFilePath}, and then workbook is closed.</p>
   * 
   * @param destFilePath destFilePath
   * @param data dataList
   */
  public void write(@RequireNonnull String templateFilePath, @RequireNonnull String destFilePath,
      @RequireNonnull List<List<T>> data) throws Exception {
    ObjectsUtil.paramRequireNonNull(templateFilePath);
    ObjectsUtil.paramRequireNonNull(destFilePath);

    try (Workbook workbook = openForWrite(templateFilePath);
        FileOutputStream out = new FileOutputStream(destFilePath);) {

      headerCheck(workbook);

      writeTableValues(workbook, data);

      workbook.write(out);
    }
  }

  /**
   * Writes table data to the designated excel file.
   * 
   * <p>{@code data} is stored to {@code workbook} created from {@code templateFilePath}, 
   *     and the method returns {@code workbook}.</p>
   * 
   * @param templateFilePath templateFilePath
   * @param data data
   */
  public Workbook write(@RequireNonnull String templateFilePath, @RequireNonnull List<List<T>> data)
      throws Exception {

    try (Workbook workbook = openForWrite(templateFilePath);) {

      headerCheck(workbook);

      writeTableValues(workbook, data);

      return workbook;
    }
  }

  /**
   * Writes table data to the designated excel file.
   * 
   * <p>{@code data} is stored to {@code workbook}.</p>
   * 
   * @param workbook workbook
   * @param data dataList
   */
  public void write(@RequireNonnull Workbook workbook, @RequireNonnull List<List<T>> data)
      throws Exception {

    headerCheck(workbook);

    writeTableValues(workbook, data);
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
    return filePath == null ? new XSSFWorkbook()
        : WorkbookFactory.create(new File(filePath), null, false);
  }

  /**
   * Obtains header list from the file at {@code templateFilePath}.
   * 
   * @param workbook workbook.
   * @throws IOException IOException
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  protected abstract void headerCheck(@RequireNonnull Workbook workbook)
      throws EncryptedDocumentException, AppException, IOException;

  private void writeTableValues(@RequireNonnull Workbook excel, @RequireNonnull List<List<T>> data)
      throws FileNotFoundException, IOException, BizLogicAppException {

    detailLog.debug(LogUtil.PARTITION_LARGE);
    detailLog.debug("starting to write excel file.");
    detailLog.debug("sheet name :" + getSheetName());

    Sheet sheet = excel.getSheet(getSheetName());

    if (sheet == null) {
      throw new BizLogicAppException("MSG_ERR_SHEET_NOT_EXIST", getSheetName());
    }

    int poiBasisTableStartColumnNumber = getPoiBasisDeterminedTableStartColumnNumber();
    int poiBasisTableStartRowNumber = getPoiBasisDeterminedTableStartRowNumber(sheet);

    // Skip the header line if the writer is OneLineHeaderFormat
    if (this instanceof IfFormatOneLineHeaderExcelTable) {
      poiBasisTableStartRowNumber++;
    }

    for (int rowNum = poiBasisTableStartRowNumber; rowNum < poiBasisTableStartRowNumber
        + data.size(); rowNum++) {

      List<T> list = data.get(rowNum - poiBasisTableStartRowNumber);

      if (sheet.getRow(rowNum) == null) {
        sheet.createRow(rowNum);
      }

      Row row = sheet.getRow(rowNum);

      for (int colNum = poiBasisTableStartColumnNumber; colNum < poiBasisTableStartColumnNumber
          + data.get(0).size(); colNum++) {

        T sourceCellData = list.get(colNum - poiBasisTableStartColumnNumber);

        if (row.getCell(colNum) == null) {
          row.createCell(colNum);
        }

        Cell destCell = row.getCell(colNum);

        writeToCell(sourceCellData, destCell);
      }
    }
  }
}
