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
package jp.ecuacion.util.excel.table.writer;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.excel.table.ExcelTable;
import jp.ecuacion.util.excel.table.IfExcelTable;
import jp.ecuacion.util.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;


/**
 * Is a parent of excel table writer classes.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTableWriter<T> extends ExcelTable<T> implements IfExcelTableWriter<T> {

  /**
   * Constructs a new instance with only the sheet name.
   *
   * <p>Defaults: {@code tableStartRowNumber = null}, {@code tableStartColumnNumber = 1}.</p>
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   */
  protected ExcelTableWriter(String sheetName) {
    super(sheetName);
  }

  /**
   * Constructs a new instance with the sheet name, the position of the excel table.
   *
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   */
  public ExcelTableWriter(String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);
  }

  /**
   * Writes table data to the specified excel file.
   *
   * <p>{@code data} is written to {@code destFilePath}, and then workbook is closed.</p>
   *
   * @param templateFilePath templateFilePath
   * @param destFilePath destFilePath
   * @param data dataList
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public void write(String templateFilePath, String destFilePath, List<List<T>> data)
      throws EncryptedDocumentException, IOException {
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
   *     and the method returns {@code workbook}.<br>
   *     The caller is responsible for closing the returned {@code workbook}.</p>
   *
   * @param templateFilePath templateFilePath
   * @param data data
   * @return workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public Workbook write(String templateFilePath, List<List<T>> data)
      throws EncryptedDocumentException, IOException {
    Workbook workbook = ExcelWriteUtil.openForWrite(templateFilePath);

    headerCheck(workbook);
    writeTableValues(workbook, data);

    return workbook;
  }

  /**
   * Writes table data to the designated excel file.
   *
   * <p>{@code data} is stored to {@code workbook}.</p>
   *
   * @param workbook workbook
   * @param data dataList
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public void write(Workbook workbook, List<List<T>> data)
      throws EncryptedDocumentException, IOException {

    headerCheck(workbook);

    writeTableValues(workbook, data);
  }

  /**
   * Provides a {@link IterableWriter} that writes rows one by one to the workbook.
   *
   * <p>The caller owns the {@code workbook} and is responsible for saving and closing it.
   *     Calling {@code close()} on the returned {@link IterableWriter} is a no-op.</p>
   *
   * @param workbook workbook
   * @return iterable writer
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public IterableWriter<T> getIterable(Workbook workbook)
      throws EncryptedDocumentException, IOException {
    headerCheck(workbook);

    ContextContainer context = ExcelWriteUtil.getReadyToWriteTableData(this, workbook,
        getSheetName(), tableStartColumnNumber);

    return new IterableWriter<T>(this, context, getNumberOfHeaderLines());
  }

  /**
   * Provides a {@link IterableWriter} that writes from {@code templateFilePath} and saves to
   *     {@code destFilePath} on close.
   *
   * <p>The returned {@link IterableWriter} owns the workbook opened from
   *     {@code templateFilePath}. Its {@link IterableWriter#close()} saves the workbook
   *     to {@code destFilePath} and closes it. Use try-with-resources to ensure
   *     the workbook is saved and closed.</p>
   *
   * @param templateFilePath templateFilePath
   * @param destFilePath destFilePath
   * @return iterable writer that owns the workbook
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public IterableWriter<T> getIterable(String templateFilePath, String destFilePath)
      throws EncryptedDocumentException, IOException {
    ObjectsUtil.requireNonNull(templateFilePath);
    ObjectsUtil.requireNonNull(destFilePath);

    Workbook workbook = ExcelWriteUtil.openForWrite(templateFilePath);
    boolean ownershipTransferred = false;
    try {
      headerCheck(workbook);

      ContextContainer context = ExcelWriteUtil.getReadyToWriteTableData(this, workbook,
          getSheetName(), tableStartColumnNumber);

      IterableWriter<T> result = new IterableWriter<T>(this, context, getNumberOfHeaderLines(),
          workbook, destFilePath);
      ownershipTransferred = true;
      return result;
    } finally {
      if (!ownershipTransferred) {
        workbook.close();
      }
    }
  }

  /**
   * Obtains header list from the file at {@code templateFilePath}.
   * 
   * @param workbook workbook.
   * @throws IOException IOException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  protected abstract void headerCheck(Workbook workbook)
      throws EncryptedDocumentException, IOException;

  private void writeTableValues(Workbook workbook, List<List<T>> data) {

    ContextContainer context = ExcelWriteUtil.getReadyToWriteTableData(this, workbook,
        getSheetName(), tableStartColumnNumber);

    final int startRowNumber = context.poiBasisTableStartRowNumber + getNumberOfHeaderLines();
    for (int rowNumber = startRowNumber; rowNumber < startRowNumber + data.size(); rowNumber++) {
      List<T> list = data.get(rowNumber - startRowNumber);
      ExcelWriteUtil.writeTableLine(this, context, rowNumber, list);
    }
  }

  /**
   * Sets {@code tableStartRowNumber} and returns {@code this} for method chaining.
   *
   * @param value See {@link ExcelTable#tableStartRowNumber}.
   * @return this writer
   */
  public ExcelTableWriter<T> tableStartRowNumber(@Nullable Integer value) {
    this.tableStartRowNumber = value;
    return this;
  }

  /**
   * Sets {@code tableStartColumnNumber} and returns {@code this} for method chaining.
   *
   * @param value See {@link ExcelTable#tableStartColumnNumber}.
   * @return this writer
   */
  public ExcelTableWriter<T> tableStartColumnNumber(int value) {
    this.tableStartColumnNumber = value;
    return this;
  }

  @SuppressWarnings("InlineMeSuggester")
  @Override
  @Deprecated
  public ExcelTableWriter<T> ignoresAdditionalColumnsOfHeaderData(boolean value) {
    return withIgnoresAdditionalColumnsOfHeaderData(value);
  }

  @Override
  public ExcelTableWriter<T> withIgnoresAdditionalColumnsOfHeaderData(boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @SuppressWarnings("InlineMeSuggester")
  @Override
  @Deprecated
  public ExcelTableWriter<T> isVerticalAndHorizontalOpposite(boolean value) {
    return withVerticalAndHorizontalOpposite(value);
  }

  @Override
  public ExcelTableWriter<T> withVerticalAndHorizontalOpposite(boolean value) {
    this.isVerticalAndHorizontalOpposite = value;
    return this;
  }

  /**
   * Writes rows one by one to the workbook.
   *
   * <p>Obtain an instance via {@link ExcelTableWriter#getIterable(Workbook)} or
   *     {@link ExcelTableWriter#getIterable(String, String)}.</p>
   *
   * <p>When constructed with an {@code ownedWorkbook}, {@link #close()} saves the workbook
   *     to {@code destPath} (if non-null) and then closes it.
   *     When constructed without one, {@code close()} is a no-op (the caller owns
   *     the workbook).</p>
   */
  public static class IterableWriter<T> implements AutoCloseable {

    private ExcelTableWriter<T> writer;
    private ContextContainer context;
    private int rowNumber;
    private @Nullable Workbook ownedWorkbook;
    private @Nullable String destPath;

    /**
     * Constructs a new instance.
     *
     * @param writer writer
     * @param context context
     * @param numberOfHeaderLines numberOfHeaderLines
     */
    public IterableWriter(ExcelTableWriter<T> writer, ContextContainer context,
        int numberOfHeaderLines) {
      this(writer, context, numberOfHeaderLines, null, null);
    }

    /**
     * Constructs a new instance with an owned workbook to be saved and closed by
     *     {@link #close()}.
     *
     * @param writer writer
     * @param context context
     * @param numberOfHeaderLines numberOfHeaderLines
     * @param ownedWorkbook the workbook this iterable owns; {@code null} means the caller
     *     owns it and {@link #close()} is a no-op
     * @param destPath the file path to save the workbook to on close;
     *     {@code null} skips saving
     */
    public IterableWriter(ExcelTableWriter<T> writer, ContextContainer context,
        int numberOfHeaderLines, @Nullable Workbook ownedWorkbook, @Nullable String destPath) {
      this.writer = writer;
      this.context = context;
      this.rowNumber = context.poiBasisTableStartRowNumber + numberOfHeaderLines;
      this.ownedWorkbook = ownedWorkbook;
      this.destPath = destPath;
    }

    /**
     * Writes one row.
     *
     * @param columnList columnList
     */
    public void write(List<T> columnList) {
      ExcelWriteUtil.writeTableLine(writer, context, rowNumber, columnList);
      rowNumber++;
    }

    @Override
    public void close() throws IOException {
      if (ownedWorkbook == null) {
        return;
      }
      
      try {
        if (destPath != null) {
          try (FileOutputStream out = new FileOutputStream(destPath)) {
            Objects.requireNonNull(ownedWorkbook).write(out);
          }
        }
      } finally {
        Objects.requireNonNull(ownedWorkbook).close();
      }
    }
  }
}
