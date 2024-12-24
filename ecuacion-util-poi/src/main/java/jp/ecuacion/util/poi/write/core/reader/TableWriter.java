package jp.ecuacion.util.poi.write.core.reader;

import jakarta.annotation.Nonnull;
import jakarta.annotation.Nullable;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.util.ObjectsUtil;

/**
 * 
 */
public class TableWriter {

  @Nonnull
  private String sheetName;

  private Integer tableStartRowNumber;
  private int tableStartColumnNumber;
  private Integer tableRowSize;
  private Integer tableColumnSize;

  /**
   * 
   * @param sheetName
   * @param tableStartRowNumber
   * @param tableStartColumnNumber
   * @param tableColumnSize
   */
  public TableWriter(@RequireNonnull String sheetName, @Nullable Integer tableStartRowNumber,
      int tableStartColumnNumber) {
    this.sheetName = ObjectsUtil.paramRequireNonNull(sheetName);
    this.tableStartRowNumber = tableStartRowNumber;
    this.tableStartColumnNumber = tableStartColumnNumber;
  }
  
  /**
   * Returns the excel sheet name the reader reads.
   * 
   * @return the sheet name
   */
  public @Nonnull String getSheetName() {
    return ObjectsUtil.returnRequireNonNull(sheetName);
  }
  

  public int getPoiBasisTableStartRowNumber() {
    return tableStartRowNumber - 1;
  }

  public int getPoiBasisTableStartColumnNumber() {
    return tableStartColumnNumber - 1;
  }
}
