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

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;


/**
 * Is a parent of excel table writer classes.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTableWriter<T> extends ExcelTable<T> implements IfExcelTableWriter<T> {

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
    ObjectsUtil.requireNonNull(templateFilePath);
    ObjectsUtil.requireNonNull(destFilePath);

    try (Workbook workbook = ExcelWriteUtil.openForWrite(templateFilePath);
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

    try (Workbook workbook = ExcelWriteUtil.openForWrite(templateFilePath);) {

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
   * Provides an {@code Iterable} writer.
   * 
   * @param workbook workbook
   */
  @Nonnull
  public IterableWriter<T> getIterable(@RequireNonnull Workbook workbook)
      throws EncryptedDocumentException, AppException, IOException {
    // Header check first, and then iterating data.
    headerCheck(workbook);

    // get the IteratorWriter
    ContextContainer context = ExcelWriteUtil.getReadyToWriteTableData(this, workbook,
        getSheetName(), tableStartColumnNumber);

    return new IterableWriter<T>(this, context, getNumberOfHeaderLines());
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

  private void writeTableValues(@RequireNonnull Workbook workbook,
      @RequireNonnull List<List<T>> data)
      throws FileNotFoundException, IOException, ExcelAppException {

    ContextContainer context = ExcelWriteUtil.getReadyToWriteTableData(this, workbook,
        getSheetName(), tableStartColumnNumber);

    final int startRowNumber = context.poiBasisTableStartRowNumber + getNumberOfHeaderLines();
    for (int rowNumber = startRowNumber; rowNumber < startRowNumber + data.size(); rowNumber++) {
      List<T> list = data.get(rowNumber - startRowNumber);
      ExcelWriteUtil.writeTableLine(this, context, rowNumber, list);
    }
  }

  @Override
  public ExcelTableWriter<T> ignoresAdditionalColumnsOfHeaderData(boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @Override
  public ExcelTableWriter<T> isVerticalAndHorizontalOpposite(boolean value) {
    this.isVerticalAndHorizontalOpposite = value;
    return this;
  }

  /**
   * Provides {@code Iterable}.
   */
  public static class IterableWriter<T> {

    private ExcelTableWriter<T> writer;
    private ContextContainer context;
    private int rowNumber;

    /**
     * Constructs a new instance.
     */
    public IterableWriter(ExcelTableWriter<T> writer, ContextContainer context,
        int numberOfheaderLines) {
      this.writer = writer;
      this.context = context;
      this.rowNumber = context.poiBasisTableStartRowNumber + numberOfheaderLines;
    }

    /**
     * Writes one line.
     * 
     * @param columnList columnList
     */
    public void write(List<T> columnList) {

      ExcelWriteUtil.writeTableLine(writer, context, rowNumber, columnList);

      rowNumber++;
    }
  }
}
