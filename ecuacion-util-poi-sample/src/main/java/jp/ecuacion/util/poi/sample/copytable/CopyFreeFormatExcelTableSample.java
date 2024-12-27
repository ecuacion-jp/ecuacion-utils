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
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.reader.cell.CellFreeFormatExcelTableReader;
import jp.ecuacion.util.poi.excel.table.writer.cell.CellFreeFormatExcelTableWriter;
import org.apache.poi.ss.usermodel.Cell;

public class CopyFreeFormatExcelTableSample {

  private static final int HEADER_START_ROW = 3;
  private static final int START_COL = 2;

  private static Path destPath;

  public static void main(String[] args) throws Exception {
    
    System.out.println("Procedure start: " + LocalDateTime.now());
    
    // read
    List<List<Cell>> dataList = read();

    // write
    write(dataList);

    System.out.println("A new excel file created and table data copied to: " + destPath.toString());

    System.out.println("Procedure finsh: " + LocalDateTime.now());
}

  private static List<List<Cell>> read() throws Exception {

    // Get the path of the excel file.
    URL sourceUrl = CopyFreeFormatExcelTableSample.class.getClassLoader().getResource("sample.xlsx");
    Path sourcePath = Path.of(sourceUrl.toURI()).toAbsolutePath();

    // Get the table data.
    return new CellFreeFormatExcelTableReader("Member", HEADER_START_ROW, START_COL, null, null)
        .read(sourcePath.toString());
  }

  private static void write(List<List<Cell>> dataList) throws Exception {

    destPath = Path.of(Paths.get("").toAbsolutePath().toString() + "/" + "result.xlsx");

    // If the created file already exists, delete it.
    if (new File(destPath.toString()).exists()) {
      Files.delete(destPath);
    }

    // Write the table data.
    new CellFreeFormatExcelTableWriter("Sheet1", HEADER_START_ROW, START_COL).write(
        null, destPath.toAbsolutePath().toString(), dataList);
  }
}
