package jp.ecuacion.util.poi.sample.copytable;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import jp.ecuacion.util.poi.read.cell.reader.CellFixedTableReader;
import jp.ecuacion.util.poi.write.cell.writer.CellFixedTableWriter;
import org.apache.poi.ss.usermodel.Cell;

public class Main {
  
  private static final String[] headerLabels =
      new String[] {"ID", "name", "date of birth", "age", "nationality"};

  private static final int HEADER_START_ROW = 3;
  private static final int START_COL = 2;
  
  public static void main(String[] args) throws Exception {
    // read
    List<List<Cell>> dataList = read();
    
    // write
    write(dataList);
  }

  private static List<List<Cell>> read() throws Exception {
    // Get the path of the excel file.
    Path sourcePath =
        Path.of(Main.class.getClassLoader().getResource("sample.xlsx").toURI()).toAbsolutePath();

    // Get the table data.
    return new CellFixedTableReader("Member", headerLabels, HEADER_START_ROW, START_COL, null)
        .getAndValidateTableValues(sourcePath.toString());
  }

  private static void write(List<List<Cell>> dataList) throws Exception {
    // Write to a new excel file. The new file is created at the current path of the execution
    // environment.

    Path newFilePath = Path.of(Main.class.getClassLoader().getResource("newFile.xlsx").toURI());
    Path destFilePath = Path.of(Paths.get("").toAbsolutePath().toString() + "/" + "copyResult.xlsx");

    // If the created file already exists, delete it.
    if (new File(destFilePath.toString()).exists()) {
      Files.delete(destFilePath);
    }

    Files.copy(newFilePath, destFilePath);
    
    new CellFixedTableWriter("Sheet1", headerLabels, HEADER_START_ROW, START_COL).write(
        newFilePath.toAbsolutePath().toString(), destFilePath.toAbsolutePath().toString(), dataList);
    
    System.out.println("A new excel file created and table data copied to: " + destFilePath.toString());
  }
}
