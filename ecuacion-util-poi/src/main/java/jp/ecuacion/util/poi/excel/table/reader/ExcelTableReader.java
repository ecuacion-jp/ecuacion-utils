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
package jp.ecuacion.util.poi.excel.table.reader;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.ValidatorFactory;
import jakarta.validation.constraints.Min;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import jp.ecuacion.lib.core.constant.EclibCoreConstants;
import jp.ecuacion.lib.core.logging.DetailLogger;
import jp.ecuacion.lib.core.util.ObjectsUtil;
import jp.ecuacion.util.poi.excel.exception.ExcelAppException;
import jp.ecuacion.util.poi.excel.exception.LoopBreakException;
import jp.ecuacion.util.poi.excel.table.ExcelTable;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspecify.annotations.Nullable;

/**
 * Is a parent of excel table reader classes.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class ExcelTableReader<T> extends ExcelTable<T> implements IfExcelTableReader<T> {

  private static DetailLogger detailLog = new DetailLogger(ExcelTableReader.class);

  /**
   * Is the row size of the table. 
   * 
   * <p>It's equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableRowSize} is 
   *     decided for the program to find an empty row.<br>
   *     When the table has a header, the row size includes the header line,
   */
  @Min(1)
  protected @Nullable Integer tableRowSizeGivenByConstructor;

  /**
   * Is the column size of the table.
   *
   * <p>It's equal to or greater than {@code 1}. <br>
   *     {@code 0} or the number smaller than that is not acceptable.<br>
   *     {@code null} is acceptable, which means {@code tableColumnSize} is
   *     decided by the length of the header.
   *     Empty header cell means it's the end of the header.<br>
   *     When the table has a header, the row size includes the header line,   */
  @Min(1)
  protected @Nullable Integer tableColumnSizeGivenByConstructor;

  /**
   * Constructs a new instance with the sheet name, the position and the size of the excel table.
   * 
   * @param sheetName See {@link ExcelTable#sheetName}.
   * @param tableStartRowNumber See {@link ExcelTable#tableStartRowNumber}.
   * @param tableStartColumnNumber See {@link ExcelTable#tableStartColumnNumber}.
   * @param tableRowSize See {@link ExcelTableReader#tableRowSizeGivenByConstructor}.
   * @param tableColumnSize See {@link ExcelTableReader#tableColumnSizeGivenByConstructor}.
   */
  public ExcelTableReader(String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber, @Nullable Integer tableRowSize,
      @Nullable Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber);

    this.tableRowSizeGivenByConstructor = tableRowSize;
    this.tableColumnSizeGivenByConstructor = tableColumnSize;

    try (ValidatorFactory factory = Validation.buildDefaultValidatorFactory()) {
      Set<ConstraintViolation<ExcelTableReader<T>>> violationSet =
          factory.getValidator().validate(this);
      if (!violationSet.isEmpty()) {
        throw new RuntimeException("Validation failed at TableReader constructor.");
      }
    }
  }

  /**
   * Reads a table data in an excel file at {@code excelPath} 
   *     and Return it in the form of {@code List<List<T>>}.
   * 
   * <p>The internal {@code List<T>} stores data in one line.<br>
   * The external {@code List} stores lines of {@code List<T>}.</p>
   *
   * @param filePath filePath
   * @throws IOException IOException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  public List<List<T>> read(String filePath)
      throws EncryptedDocumentException, IOException {
    ObjectsUtil.requireNonNull(filePath);

    try (Workbook excel = ExcelReadUtil.openForRead(filePath);) {
      return read(excel);
    }
  }

  /**
   * Reads a table data from {@code workbook}
   *     and returns it in the form of {@code List<List<T>>}.
   *
   * <p>The internal {@code List<T>} stores data in one line.<br>
   * The external {@code List} stores lines of {@code List<T>}.</p>
   *
   * @param workbook workbook
   * @throws IOException IOException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  public List<List<T>> read(Workbook workbook)
      throws EncryptedDocumentException, IOException {

    // validate the header line
    List<List<T>> headerData = readTableData(workbook, true);
    validateHeaderData(headerData);

    // obtain data
    List<List<T>> rtnData = readTableData(workbook);
    updateAndGetHeaderData(rtnData);

    return rtnData;
  }

  /**
   * Provides an {@code Iterable} reader over the data rows of the table.
   *
   * @param workbook workbook
   * @throws IOException IOException
   * @throws EncryptedDocumentException EncryptedDocumentException
   */
  public Iterable<List<T>> getIterable(Workbook workbook)
      throws EncryptedDocumentException, IOException {

    // validate the header line
    List<List<T>> headerData = readTableData(workbook, true);
    validateHeaderData(headerData);

    // obtain data
    List<List<T>> rtnData = readTableData(workbook);
    updateAndGetHeaderData(rtnData);

    // get the IteratorReader
    ContextContainer context = getReadyToReadTableData(this, workbook, getSheetName(),
        tableStartColumnNumber, null, false);

    return new IterableReader<T>(this, context, getNumberOfHeaderLines());
  }

  /*
   * get Table Values in the form of the list of the lists.
   */
  private List<List<T>> readTableData(Workbook workbook) {
    return readTableData(workbook, false);
  }

  /*
   * get Table Values in the form of the list of the lists.
   */
  private List<List<T>> readTableData(Workbook workbook, boolean readsHeaderOnly) {

    // when readsHeaderOnly == true, return data is used to validate the header labels,
    // so ignoresColumnSizeSetInReader should also be true.
    ContextContainer context =
        getReadyToReadTableData(this, workbook, getSheetName(), tableStartColumnNumber,
            (readsHeaderOnly) ? getNumberOfHeaderLines() : null, readsHeaderOnly);

    List<List<T>> rowList = new ArrayList<>();
    try {
      for (int rowNumber =
          context.poiBasisTableStartRowNumber; rowNumber <= ContextContainer.max; rowNumber++) {
        List<T> colList = readTableLine(this, context, rowNumber);
        rowList.add(colList);
      }
    } catch (LoopBreakException ex) {
      // do nothing, just finish the loop.
    }

    detailLog.debug("finishing to read excel file. sheet name :" + getSheetName());
    detailLog.debug(EclibCoreConstants.PARTITION_LARGE);

    return rowList;
  }

  /**
   * Returns tableRowSize, may be {@code null}. 
   */
  public @Nullable Integer getTableRowSize() {
    return tableRowSizeGivenByConstructor;
  }

  /**
   * Returns tableColumnSize.
   *
   * @param sheet sheet
   * @param poiBasisDeterminedTableStartRowNumber poiBasisDeterminedTableStartRowNumber
   * @param poiBasisDeterminedTableStartColumnNumber poiBasisDeterminedTableStartColumnNumber
   * @param ignoresColumnSizeSetInReader ignoresColumnSizeSetInReader
   * @throws ExcelAppException ExcelAppException
   */
  public Integer getTableColumnSize(Sheet sheet, int poiBasisDeterminedTableStartRowNumber,
      int poiBasisDeterminedTableStartColumnNumber, boolean ignoresColumnSizeSetInReader)
      throws ExcelAppException {
    ObjectsUtil.requireNonNull(sheet);

    if (tableColumnSizeGivenByConstructor != null && !ignoresColumnSizeSetInReader) {
      return tableColumnSizeGivenByConstructor;
    }

    // the following is executed when tableColumnSize value needs to be analyzed dynamically.
    int columnNumber = poiBasisDeterminedTableStartColumnNumber;
    Cell cell;
    while (true) {
      if (isVerticalAndHorizontalOpposite) {
        Row row = sheet.getRow(columnNumber);
        // If the cell is null, that means header is end.
        if (row == null || row.getCell(poiBasisDeterminedTableStartRowNumber) == null) {
          break;
        }

        cell = row.getCell(poiBasisDeterminedTableStartRowNumber);

      } else {
        Row row = sheet.getRow(poiBasisDeterminedTableStartRowNumber);
        // If the cell is null, that means header is end.
        if (row == null || row.getCell(columnNumber) == null) {
          break;
        }

        cell = row.getCell(columnNumber);
      }

      if (isCellDataEmpty(getCellData(cell, tableStartColumnNumber + columnNumber + 1))) {
        break;
      }

      columnNumber++;
    }

    int size = columnNumber - poiBasisDeterminedTableStartColumnNumber;

    if (size == 0) {
      throw new ExcelAppException("jp.ecuacion.util.poi.excel.reader.ColumnSizeIsZero.message",
          sheet.getSheetName(), Integer.toString(poiBasisDeterminedTableStartRowNumber + 1),
          Integer.toString(poiBasisDeterminedTableStartColumnNumber + 1)).sheet(sheet);
    }

    return size;
  }

  /**
   * sets {@code tableColumnSize}.
   * 
   * <p>tableColumnSize can be set by the header length, 
   *     but the instance method cannot be called from constructors so the setter is needed.</p>
   *     
   * <p>This method set the final value of the column size, 
   *     so the argument is not {@code Integer}, but primitive {@code int}.
   * 
   * @param tableColumnSize tableColumnSize.
   */
  public void setTableColumnSize(int tableColumnSize) {
    this.tableColumnSizeGivenByConstructor = tableColumnSize;
  }

  /**
   * Provides common procedure for reading one line of a table.
   *
   * <p>It's called from both {@code ExcelTableReader} and {@code IteratorReader},
   *     so it is defined as a static method.</p>
   *
   * @param reader reader
   * @param context context
   * @param rowNumber rowNumber
   * @throws ExcelAppException ExcelAppException
   */
  static <T> List<T> readTableLine(ExcelTableReader<T> reader, ContextContainer context,
      int rowNumber) throws ExcelAppException {
    detailLog.debug(EclibCoreConstants.PARTITION_MEDIUM);
    detailLog.debug("row number：" + rowNumber);

    if (rowNumber == ContextContainer.max) {
      throw new RuntimeException("'max':" + ContextContainer.max + " exceeded.");
    }

    if (context.tableRowSize != null
        && rowNumber >= context.poiBasisTableStartRowNumber + context.tableRowSize) {
      throw new LoopBreakException();
    }

    List<T> colList = new ArrayList<>();
    boolean isEmptyRow = true;

    int tableColumnSize = java.util.Objects.requireNonNull(context.tableColumnSize);
    for (int j = context.poiBasisTableStartColumnNumber;
        j < context.poiBasisTableStartColumnNumber + tableColumnSize; j++) {

      if (reader.isVerticalAndHorizontalOpposite()) {
        Row row = context.sheet.getRow(j);
        if (row == null || row.getCell(rowNumber) == null) {
          colList.add(null);

        } else {
          Cell cell = row.getCell(rowNumber);
          T cellData = reader.getCellData(cell, j + 1);
          colList.add(cellData);
        }

      } else {
        Row row = context.sheet.getRow(rowNumber);
        if (row == null || row.getCell(j) == null) {
          colList.add(null);

        } else {
          Cell cell = row.getCell(j);
          T cellData = reader.getCellData(cell, j + 1);
          colList.add(cellData);
        }
      }
    }

    for (T colData : colList) {
      if (!reader.isCellDataEmpty(colData)) {
        isEmptyRow = false;
        break;
      }
    }

    if (isEmptyRow) {
      detailLog.debug("(no data in the line)");
      detailLog.debug(EclibCoreConstants.PARTITION_MEDIUM);

      if (context.tableRowSize == null) {
        throw new LoopBreakException();

      } else {
        // An empty row within a fixed row size is represented as an empty list.
        return new ArrayList<>();
      }
    }

    return colList;
  }

  /**
   * Gets ready to read table data.
   * 
   * @param ignoresColumnSizeSetInReader It is {@code true} means 
   *     that even if the reader determines the column size,
   *     this method obtains all the columns as long as the header column exists.
   */
  public static <T> ContextContainer getReadyToReadTableData(ExcelTableReader<T> reader,
      Workbook workbook, String sheetName, int tableStartColumnNumber,
      @Nullable Integer numberOfHeaderLinesIfReadsHeaderOnlyOrNull,
      boolean ignoresColumnSizeSetInReader)
      throws ExcelAppException {
    detailLog.debug(EclibCoreConstants.PARTITION_LARGE);
    detailLog.debug("starting to read excel file.");
    detailLog.debug("sheet name :" + sheetName);

    Sheet sheet = workbook.getSheet(sheetName);

    if (sheet == null) {
      throw new ExcelAppException("jp.ecuacion.util.poi.excel.SheetNotExist.message", sheetName);
    }

    Integer tableRowSize =
        numberOfHeaderLinesIfReadsHeaderOnlyOrNull == null ? reader.getTableRowSize()
            : numberOfHeaderLinesIfReadsHeaderOnlyOrNull;

    // poiBasis means the top-left position is (0, 0)
    // while tableStartRowNumber / tableStartColumnNumber >= 1.
    final int poiBasisTableStartRowNumber =
        reader.getPoiBasisDeterminedTableStartRowNumber(sheet, tableStartColumnNumber);
    final int poiBasisTableStartColumnNumber = reader.getPoiBasisDeterminedTableStartColumnNumber();
    ContextContainer context =
        new ContextContainer(sheet, poiBasisTableStartRowNumber, poiBasisTableStartColumnNumber,
            tableRowSize, reader.getTableColumnSize(sheet, poiBasisTableStartRowNumber,
                poiBasisTableStartColumnNumber, ignoresColumnSizeSetInReader));

    return context;
  }

  /**
   * Provides {@code Iterable}.
   */
  public static class IterableReader<T> implements Iterable<List<T>> {

    private IteratorReader<T> iterator;

    /**
     * Constructs a new instance.
     */
    public IterableReader(ExcelTableReader<T> reader, ContextContainer context,
        int numberOfHeaderLines) {
      this.iterator = new IteratorReader<T>(reader, context, numberOfHeaderLines);
    }

    @Override
    public Iterator<List<T>> iterator() {
      return iterator;
    }
  }

  /**
   * Provides Iterator.
   * 
  * @param <T> See {@link IfExcelTable}.
   */
  public static class IteratorReader<T> implements Iterator<List<T>> {

    private ExcelTableReader<T> reader;
    private ContextContainer context;
    private boolean hasNext = true;
    private int rowNumber;

    /**
     * Constructs a new instance.
     */
    public IteratorReader(ExcelTableReader<T> reader, ContextContainer context,
        int numberOfHeaderLines) {
      this.reader = reader;
      this.context = context;
      this.rowNumber = context.poiBasisTableStartRowNumber + numberOfHeaderLines;
    }

    @Override
    public boolean hasNext() {
      return hasNext;
    }

    @Override
    public List<T> next() {
      if (!hasNext) {
        throw new NoSuchElementException();
      }

      List<T> rtn = null;

      try {
        rtn = readTableLine(reader, context, rowNumber);

        rowNumber++;

        // check for hasNext
        try {
          readTableLine(reader, context, rowNumber);
        } catch (LoopBreakException ex) {
          hasNext = false;
        }

        return rtn;

      } catch (ExcelAppException ex) {
        throw ex;
      }

    }
  }

  @Override
  public ExcelTableReader<T> ignoresAdditionalColumnsOfHeaderData(boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }

  @Override
  public ExcelTableReader<T> isVerticalAndHorizontalOpposite(boolean value) {
    this.isVerticalAndHorizontalOpposite = value;
    return this;
  }
}
