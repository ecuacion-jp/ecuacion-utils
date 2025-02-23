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
import jp.ecuacion.lib.core.util.ExceptionUtil;
import jp.ecuacion.lib.core.util.LogUtil;
import jp.ecuacion.lib.core.util.PropertyFileUtil.Arg;
import jp.ecuacion.util.poi.excel.table.ExcelTable.ContextContainer;
import jp.ecuacion.util.poi.excel.table.IfFormatOneLineHeaderExcelTable;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment.WorkbookNotFoundException;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.formula.eval.NotImplementedFunctionException;
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
  private ExceptionUtil exUtil = new ExceptionUtil();

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
      throw new BizLogicAppException("jp.ecuacion.util.poi.excel.SheetNotExist.message", sheetName);
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

  /**
   * Catches {@code Exception}s which are thrown 
   *     when {@code workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()} is called
   *     and changes it to a {@code BizLogicAppException} with an appropriate message.
   * 
   * <p>When an excel file is created and uploaded by users, 
   *     {@code Exception}s according to the content of the file 
   *     should be understandable to the users.</p>
   * 
   * @param workbook workbook
   * @throws BizLogicAppException BizLogicAppException
   */
  public void evaluateAllFormulas(Workbook workbook) throws BizLogicAppException {
    evaluateAllFormulas(workbook, null);
  }

  /**
   * Catches {@code Exception}s which are thrown 
   *     when {@code workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()} is called
   *     and changes it to a {@code BizLogicAppException} with an appropriate message.
   * 
   * <p>When an excel file is created and uploaded by users, 
   *     {@code Exception}s according to the content of the file 
   *     should be understandable to the users.</p>
   * 
   * @param workbook workbook
   * @param fileInfo filename or file path of the excel file to add to the message
   * @throws BizLogicAppException BizLogicAppException
   */
  public void evaluateAllFormulas(Workbook workbook, String fileInfo) throws BizLogicAppException {
    Arg fileInfoArg = getFileInfoString(fileInfo);

    try {
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();

    } catch (NotImplementedException ex) {

      String sheetAndCell = ex.getMessage().replace("Error evaluating cell ", "");
      String sheet = sheetAndCell.split("!")[0];
      String cell = sheetAndCell.split("!")[1];

      Arg reason = Arg.message("jp.ecuacion.util.poi.excel.ExcelWriteUtil"
          + ".NotImplementedException.ReasonUnknown.message");

      if (ex.getCause() instanceof NotImplementedFunctionException) {
        NotImplementedFunctionException ex2 = (NotImplementedFunctionException) ex.getCause();
        String msg = "jp.ecuacion.util.poi.excel.ExcelWriteUtil.NotImplementedException."
            + "ReasonUnimplementedFunction.message";
        reason = Arg.message(msg, Arg.string(ex2.getMessage().replace("_xlfn.", "")));
      }

      throw new BizLogicAppException(
          "jp.ecuacion.util.poi.excel.ExcelWriteUtil.NotImplementedException.message",
          ArrayUtils.addAll(
              ArrayUtils.addAll(Arg.strings(sheet, cell), reason), fileInfoArg)
          );

    } catch (FormulaParseException ex) {
      throwBizLogicExceptionForUnknownException(ex, fileInfo);

    } catch (IllegalStateException ex) {
      String msg1 = ex.getMessage();

      String startsWith1 = "Failed to evaluate cell: ";
      if (!msg1.startsWith(startsWith1)) {
        // In case of unknown cause.
        throwBizLogicExceptionForUnknownException(ex, fileInfo);
      }

      String[] arg1Tmp1Arr = msg1.replace(startsWith1, "").split(",");
      String sheetame = arg1Tmp1Arr[0].split("!")[0];
      String cellName = arg1Tmp1Arr[0].split("!")[1];
      String errorOccuredFunction = arg1Tmp1Arr[1].replace("value: ", "");

      if (ex.getCause() != null && ex.getCause() instanceof IllegalStateException) {
        IllegalStateException ex2 = (IllegalStateException) ex.getCause();

        if (ex2.getCause() != null && ex2.getCause() instanceof WorkbookNotFoundException) {
          WorkbookNotFoundException ex3 = (WorkbookNotFoundException) ex2.getCause();

          String startsWith = "Could not resolve external workbook name ";
          if (ex3.getMessage().startsWith(startsWith)) {
            String arg3Tmp1 = ex3.getMessage().replace(startsWith + "'", "");
            String fileInfoInFunction = arg3Tmp1.substring(0, arg3Tmp1.indexOf("'"));

            throw new BizLogicAppException(
                "jp.ecuacion.util.poi.excel.ExcelWriteUtil.WorkbookNotFoundException.message",
                ArrayUtils.addAll(
                    Arg.strings(sheetame, cellName, errorOccuredFunction, fileInfoInFunction),
                    fileInfoArg));

          } else {
            // In case of unknown cause.
            throwBizLogicExceptionForIllegalStateExceptionFailedToEvaluateCell(ex3, fileInfo,
                sheetame, cellName, errorOccuredFunction);
          }
        } else {
          // In case of unknown cause.
          throwBizLogicExceptionForIllegalStateExceptionFailedToEvaluateCell(ex2, fileInfo,
              sheetame, cellName, errorOccuredFunction);
        }

      } else {
        // In case of unknown cause.
        throwBizLogicExceptionForIllegalStateExceptionFailedToEvaluateCell(ex, fileInfo, sheetame,
            cellName, errorOccuredFunction);
      }

    } catch (

    Exception ex) {
      throwBizLogicExceptionForUnknownException(ex, fileInfo);
    }
  }

  private void throwBizLogicExceptionForUnknownException(Exception ex, String fileInfo)
      throws BizLogicAppException {
    StringBuilder sb = new StringBuilder();
    exUtil.getExceptionListWithMessages(ex).stream()
        .forEach(e -> sb.append("  - " + e.getMessage()));
    Arg fileInfoArg = getFileInfoString(fileInfo);

    throw new BizLogicAppException(
        "jp.ecuacion.util.poi.excel.ExcelWriteUtil.DetailUnknown.message",
        ArrayUtils.addAll(new Arg[] {fileInfoArg}, Arg.string(sb.toString())));
  }

  private void throwBizLogicExceptionForIllegalStateExceptionFailedToEvaluateCell(Exception ex,
      String fileInfo, String sheetName, String cellName, String errorOccuredFunction)
      throws BizLogicAppException {
    StringBuilder sb = new StringBuilder();
    exUtil.getExceptionListWithMessages(ex).stream()
        .forEach(e -> sb.append("      - " + e.getMessage()));
    Arg fileInfoArg = getFileInfoString(fileInfo);

    throw new BizLogicAppException(
        "jp.ecuacion.util.poi.excel.ExcelWriteUtil.DetailUnknown.message",
        ArrayUtils.addAll(
            ArrayUtils.addAll(Arg.strings(sheetName, cellName, errorOccuredFunction), fileInfoArg),
            Arg.strings(sb.toString())));
  }

  private Arg getFileInfoString(String fileInfo) {

    String infoNone = "jp.ecuacion.util.poi.excel.ExcelWriteUtil.FileInfoLabel.None.message";
    Arg fileInfoLabel = fileInfo == null ? Arg.message(infoNone) : Arg.string(fileInfo);

    return fileInfoLabel;
  }
}
