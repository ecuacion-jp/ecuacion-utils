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

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import jp.ecuacion.util.excel.table.reader.concrete.CellFreeExcelTableReader;
import jp.ecuacion.util.excel.table.writer.concrete.CellFreeExcelTableWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class CopyFreeFormatExcelTableSample {

  private static final int HEADER_START_ROW = 4;
  private static final int START_COL = 2;

  private static Path destPath;

  public static void main(String[] args) throws Exception {

    Logger logger = LoggerFactory.getLogger(CopyFreeFormatExcelTableSample.class);

    logger.info("Procedure started.");

    // read
    List<List<Cell>> dataList = read();

    // write
    write(dataList);

    logger.info("A new excel file created and table data copied to: " + destPath.toString());

    logger.info("Procedure finshed.");
  }

  private static List<List<Cell>> read() throws Exception {

    // Get the table data.
    return new CellFreeExcelTableReader("Member")
        .tableStartRowNumber(HEADER_START_ROW)
        .tableStartColumnNumber(START_COL)
        .read(Path.of("test-data/sample.xlsx").toAbsolutePath().toString());
  }

  private static void write(List<List<Cell>> dataList) throws Exception {

    new File("target/test-result").mkdirs();

    String templatePath = Path.of("test-data/template.xlsx").toAbsolutePath().toString();
    destPath = Path.of("target/test-result/result.xlsx").toAbsolutePath();

    // If the created file already exists, delete it.
    if (new File(destPath.toString()).exists()) {
      Files.delete(destPath);
    }

    // Write the table data.
    new CellFreeExcelTableWriter("Sheet1")
        .tableStartRowNumber(HEADER_START_ROW)
        .tableStartColumnNumber(START_COL)
        .write(templatePath, destPath.toAbsolutePath().toString(), dataList);
  }
}
