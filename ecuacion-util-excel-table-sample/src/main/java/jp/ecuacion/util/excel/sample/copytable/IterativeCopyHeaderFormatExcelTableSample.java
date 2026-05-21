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
package jp.ecuacion.util.excel.sample.copytable;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.excel.table.reader.concrete.CellOneLineHeaderExcelTableReader;
import jp.ecuacion.util.excel.table.writer.ExcelTableWriter.IterableWriter;
import jp.ecuacion.util.excel.table.writer.concrete.CellOneLineHeaderExcelTableWriter;
import jp.ecuacion.util.excel.util.ExcelReadUtil;
import jp.ecuacion.util.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class IterativeCopyHeaderFormatExcelTableSample {

  private static final String[] headerLabels =
      new String[] {"ID", "name", "date of birth", "age", "nationality"};

  private static final int HEADER_START_ROW = 3;
  private static final int START_COL = 2;

  private static String destPath;

  public static void main(String[] args) throws Exception {

    Logger logger = LoggerFactory.getLogger(IterativeCopyHeaderFormatExcelTableSample.class);

    logger.info("Procedure started.");

    try (Workbook readWb = openToRead();
        Workbook writeWb = openToWrite();
        FileOutputStream out = openToOutput();) {

      // reader
      Iterable<List<Cell>> itReader = new CellOneLineHeaderExcelTableReader("Member", headerLabels)
          .tableStartRowNumber(HEADER_START_ROW)
          .tableStartColumnNumber(START_COL)
          .tableRowSize(3)
          .getIterable(readWb);
      // writer
      IterableWriter<Cell> itWriter =
          new CellOneLineHeaderExcelTableWriter("Sheet1", headerLabels)
              .tableStartRowNumber(HEADER_START_ROW)
              .tableStartColumnNumber(START_COL)
              .getIterable(writeWb);

      for (List<Cell> list : itReader) {
        // write
        itWriter.write(list);
      }

      ExcelWriteUtil.saveToFile(writeWb, out);
    }

    logger.info("A new excel file created and table data copied to: " + destPath.toString());
    logger.info("Procedure finished.");
  }

  private static Workbook openToRead() throws EncryptedDocumentException, IOException {
    return ExcelReadUtil.openForRead(Path.of("test-data/sample.xlsx").toAbsolutePath().toString());
  }

  private static Workbook openToWrite() throws Exception {
    return ExcelWriteUtil.openForWrite(
        Path.of("test-data/template.xlsx").toAbsolutePath().toString());
  }

  private static FileOutputStream openToOutput() throws Exception {

    Files.createDirectories(Path.of("target/test-result"));

    destPath = Path.of("target/test-result/result.xlsx").toAbsolutePath().toString();

    // If the created file already exists, delete it.
    if (Files.exists(Path.of(destPath))) {
      Files.delete(Path.of(destPath));
    }

    return ExcelWriteUtil.openForOutput(destPath);
  }
}
