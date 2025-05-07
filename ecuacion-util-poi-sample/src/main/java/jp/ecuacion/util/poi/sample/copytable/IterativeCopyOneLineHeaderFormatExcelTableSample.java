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
package jp.ecuacion.util.poi.sample.copytable;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.reader.concrete.CellOneLineHeaderExcelTableReader;
import jp.ecuacion.util.poi.excel.table.writer.ExcelTableWriter.IterableWriter;
import jp.ecuacion.util.poi.excel.table.writer.concrete.CellOneLineHeaderExcelTableWriter;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import jp.ecuacion.util.poi.excel.util.ExcelWriteUtil;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class IterativeCopyOneLineHeaderFormatExcelTableSample {

  private static final String[] headerLabels =
      new String[] {"ID", "name", "date of birth", "age", "nationality"};

  private static final int HEADER_START_ROW = 3;
  private static final int START_COL = 2;
  
  private static String destPath;

  private static ExcelReadUtil readUtil = new ExcelReadUtil();

  public static void main(String[] args) throws Exception {

    Logger logger = LoggerFactory.getLogger(IterativeCopyOneLineHeaderFormatExcelTableSample.class);

    logger.info("Procedure started.");

    try (Workbook readWb = openToRead();
        Workbook writeWb = openToWrite();
        FileOutputStream out = openToOutput();) {

      // reader
      Iterable<List<Cell>> itReader = new CellOneLineHeaderExcelTableReader("Member", headerLabels,
          HEADER_START_ROW, START_COL, 3).getIterable(readWb);
      // writer
      IterableWriter<Cell> itWriter =
          new CellOneLineHeaderExcelTableWriter("Sheet1", headerLabels, HEADER_START_ROW, START_COL)
              .getIterable(writeWb);

      for (List<Cell> list : itReader) {
        // write
        itWriter.write(list);
      }

      ExcelWriteUtil.saveToFile(writeWb, out);
    }

    logger.info("A new excel file created and table data copied to: " + destPath.toString());
    logger.info("Procedure finshed.");
  }

  private static Workbook openToRead()
      throws URISyntaxException, EncryptedDocumentException, IOException {
    // Get the path of the excel file.
    URL sourceUrl = IterativeCopyOneLineHeaderFormatExcelTableSample.class.getClassLoader()
        .getResource("sample.xlsx");
    String sourcePath = Path.of(sourceUrl.toURI()).toAbsolutePath().toString();

    return readUtil.openForRead(sourcePath);
  }

  private static Workbook openToWrite() throws Exception {

    // Create a new file from a template excel file and write the table data to it.
    // The new file will be created as "result.xlsx" at the current path.
    URL templateUrl = IterativeCopyOneLineHeaderFormatExcelTableSample.class.getClassLoader()
        .getResource("template.xlsx");
    String templatePath = Path.of(templateUrl.toURI()).toAbsolutePath().toString();

    return ExcelWriteUtil.openForWrite(templatePath);
  }

  private static FileOutputStream openToOutput() throws Exception {

    destPath = Path.of(Paths.get("").toAbsolutePath().toString() + "/" + "result.xlsx")
        .toAbsolutePath().toString();

    // If the created file already exists, delete it.
    if (new File(destPath).exists()) {
      Files.delete(Path.of(destPath));
    }

    return ExcelWriteUtil.openForOutput(destPath);
  }
}
