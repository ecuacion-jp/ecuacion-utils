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
package jp.ecuacion.util.poi.excel.table.writer.concrete;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.poi.excel.table.reader.concrete.CellOneLineHeaderExcelTableReader;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import jp.ecuacion.util.poi.excel.util.ExcelWriteUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class CellOneLineHeaderExcelTableWriterTest {

  private ExcelReadUtil readUtil = new ExcelReadUtil();
  private ExcelWriteUtil writeUtil = new ExcelWriteUtil();
  private final String origFilename = this.getClass().getSimpleName() + ".xlsx";

  private String getDestExcelFilePath(String filename) {
    String destExcelDirPath = "target/excel";
    String destExcelFilePath = destExcelDirPath + "/" + filename;
    new File(destExcelDirPath).mkdirs();

    return destExcelFilePath;
  }

  private String getDestFilename(String methodName) {
    return this.getClass().getSimpleName() + "-" + methodName + ".xlsx";
  }

  @Test
  public void normalTableTest() throws Exception {
    String destFilename = getDestFilename("normalTableTest");
    String origExcelPath = "src/test/resources/" + origFilename;
    String destExcelFilePath = getDestExcelFilePath(destFilename);
    final String[] HEADER_LABELS = new String[] {"header1", "header2", "header3"};

    List<List<Cell>> rowList =
        new CellOneLineHeaderExcelTableReader("copy-from", HEADER_LABELS, 2, 1, null)
            .read(origExcelPath);
    Workbook wb = writeUtil.openForWrite(origExcelPath);
    String copyToSheetName = "copy-to-normalTableTest";

    // try-finally added to save the tested excel file.
    try {

      // normal copy

      new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 2, 1).write(wb,
          rowList);
      Sheet sheet = wb.getSheet(copyToSheetName);

      Assertions.assertEquals("header1", readUtil.getStringFromCell(sheet.getRow(1).getCell(0)));
      Assertions.assertEquals("header2", readUtil.getStringFromCell(sheet.getRow(1).getCell(1)));
      Assertions.assertEquals("header3", readUtil.getStringFromCell(sheet.getRow(1).getCell(2)));
      Assertions.assertEquals(null, sheet.getRow(1).getCell(3));

      Assertions.assertEquals("data1-1", readUtil.getStringFromCell(sheet.getRow(2).getCell(0)));
      Assertions.assertEquals("data1-2", readUtil.getStringFromCell(sheet.getRow(2).getCell(1)));
      Assertions.assertEquals("data1-3", readUtil.getStringFromCell(sheet.getRow(2).getCell(2)));
      Assertions.assertEquals(null, sheet.getRow(2).getCell(3));

      Assertions.assertEquals("data2-1", readUtil.getStringFromCell(sheet.getRow(3).getCell(0)));
      Assertions.assertEquals("data2-2", readUtil.getStringFromCell(sheet.getRow(3).getCell(1)));
      Assertions.assertEquals("data2-3", readUtil.getStringFromCell(sheet.getRow(3).getCell(2)));
      Assertions.assertEquals(null, sheet.getRow(2).getCell(3));

      // copy to whitespace (row == null)

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 6, 1).write(wb,
            rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.reader.ColumnSizeIsZero.message",
            ex.getMessageId());
      }

      // copy to whitespace (row != null)

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 9, 1).write(wb,
            rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.reader.ColumnSizeIsZero.message",
            ex.getMessageId());
      }

      // copy to the position where is a smaller size of header labels, with additional columns
      // not allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 13, 1).write(wb,
            rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a smaller size of header labels, with additional columns
      // allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 13, 1)
            .ignoresAdditionalColumnsOfHeaderData(true).write(wb, rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a larger size of header labels, with additional columns
      // not allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 17, 1).write(wb,
            rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a larger size of header labels, with additional columns
      // allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 17, 1)
            .ignoresAdditionalColumnsOfHeaderData(true).write(wb, rowList);

      } catch (BizLogicAppException ex) {
        Assertions.fail();
      }

    } finally {
      // delete previous test data.
      if (new File(destExcelFilePath).exists()) {
        Files.delete(Path.of(destExcelFilePath));
      }

      writeUtil.saveToFile(wb, new FileOutputStream(destExcelFilePath));
    }
  }

  @Test
  public void tableWithNullCellseTest() throws Exception {
    String destFilename = getDestFilename("tableWithNullCellseTest");
    String origExcelPath = "src/test/resources/" + origFilename;
    String destExcelFilePath = getDestExcelFilePath(destFilename);
    final String[] HEADER_LABELS = new String[] {"header1", "header2", "header3"};

    List<List<Cell>> rowList =
        new CellOneLineHeaderExcelTableReader("copy-from", HEADER_LABELS, 8, 1, null)
            .read(origExcelPath);
    Workbook wb = writeUtil.openForWrite(origExcelPath);
    String copyToSheetName = "copy-to-tableWithNullCellseTest";

    // try-finally added to save the tested excel file.
    try {

      // normal copy

      new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 1).write(wb,
          rowList);
      Sheet sheet = wb.getSheet(copyToSheetName);

      Assertions.assertEquals("header1", readUtil.getStringFromCell(sheet.getRow(0).getCell(0)));
      Assertions.assertEquals("header2", readUtil.getStringFromCell(sheet.getRow(0).getCell(1)));
      Assertions.assertEquals("header3", readUtil.getStringFromCell(sheet.getRow(0).getCell(2)));

      Assertions.assertEquals("data1-1", readUtil.getStringFromCell(sheet.getRow(1).getCell(0)));
      Assertions.assertEquals("data1-2", readUtil.getStringFromCell(sheet.getRow(1).getCell(1)));
      Assertions.assertEquals(null, readUtil.getStringFromCell(sheet.getRow(1).getCell(2)));

    } finally {
      // delete previous test data.
      if (new File(destExcelFilePath).exists()) {
        Files.delete(Path.of(destExcelFilePath));
      }

      writeUtil.saveToFile(wb, new FileOutputStream(destExcelFilePath));
    }
  }

  @Test
  public void verticalHeaderTableTest() throws Exception {
    String destFilename = getDestFilename("verticalHeaderTableTest");
    String origExcelPath = "src/test/resources/" + origFilename;
    String destExcelFilePath = getDestExcelFilePath(destFilename);
    final String[] HEADER_LABELS = new String[] {"header1", "header2", "header3"};

    List<List<Cell>> rowList =
        new CellOneLineHeaderExcelTableReader("copy-from", HEADER_LABELS, 2, 1, null)
            .read(origExcelPath);
    Workbook wb = writeUtil.openForWrite(origExcelPath);
    String copyToSheetName = "copy-to-verticalHeaderTableTest";

    // try-finally added to save the tested excel file.
    try {

      // normal copy

      new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 3, 2)
          .isVerticalAndHorizontalOpposite(true).write(wb, rowList);
      Sheet sheet = wb.getSheet(copyToSheetName);

      Assertions.assertEquals("header1", readUtil.getStringFromCell(sheet.getRow(1).getCell(2)));
      Assertions.assertEquals("header2", readUtil.getStringFromCell(sheet.getRow(2).getCell(2)));
      Assertions.assertEquals("header3", readUtil.getStringFromCell(sheet.getRow(3).getCell(2)));

      Assertions.assertEquals("data1-1", readUtil.getStringFromCell(sheet.getRow(1).getCell(3)));
      Assertions.assertEquals("data1-2", readUtil.getStringFromCell(sheet.getRow(2).getCell(3)));
      Assertions.assertEquals("data1-3", readUtil.getStringFromCell(sheet.getRow(3).getCell(3)));

      Assertions.assertEquals("data2-1", readUtil.getStringFromCell(sheet.getRow(1).getCell(4)));
      Assertions.assertEquals("data2-2", readUtil.getStringFromCell(sheet.getRow(2).getCell(4)));
      Assertions.assertEquals("data2-3", readUtil.getStringFromCell(sheet.getRow(3).getCell(4)));

      // copy to whitespace

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 8)
            .isVerticalAndHorizontalOpposite(true).write(wb, rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.reader.ColumnSizeIsZero.message",
            ex.getMessageId());
      }

      // copy to the position where is a smaller size of header labels, with additional columns
      // not allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 11)
            .isVerticalAndHorizontalOpposite(true).write(wb, rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a smaller size of header labels, with additional columns
      // allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 11)
            .ignoresAdditionalColumnsOfHeaderData(true).isVerticalAndHorizontalOpposite(true)
            .write(wb, rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a larger size of header labels, with additional columns
      // not allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 16)
            .isVerticalAndHorizontalOpposite(true).write(wb, rowList);
        Assertions.fail();

      } catch (BizLogicAppException ex) {
        Assertions.assertEquals("jp.ecuacion.util.poi.excel.NumberOfTableHeadersDiffer.message",
            ex.getMessageId());
      }

      // copy to the position where is a larger size of header labels, with additional columns
      // allowed.

      try {
        new CellOneLineHeaderExcelTableWriter(copyToSheetName, HEADER_LABELS, 1, 16)
            .ignoresAdditionalColumnsOfHeaderData(true).isVerticalAndHorizontalOpposite(true)
            .write(wb, rowList);

      } catch (BizLogicAppException ex) {
        Assertions.fail();
      }

    } finally {
      // delete previous test data.
      if (new File(destExcelFilePath).exists()) {
        Files.delete(Path.of(destExcelFilePath));
      }

      writeUtil.saveToFile(wb, new FileOutputStream(destExcelFilePath));
    }
  }
}
