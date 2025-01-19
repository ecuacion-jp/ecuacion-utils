package jp.ecuacion.util.poi.excel.table.reader.concrete;

import jakarta.annotation.Nullable;
import java.util.HashMap;
import java.util.Map;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.reader.ExcelTableReader;
import jp.ecuacion.util.poi.excel.table.reader.IfDataTypeStringExcelTableReader;

/**
 * Adds String feature to {@link ExcelTableReader}.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public abstract class StringExcelTableReader extends ExcelTableReader<String>
    implements IfDataTypeStringExcelTableReader {

  protected String defaultDateTimeFormat;
  protected Map<Integer, String> columnDateTimeFormatMap = new HashMap<>();

  /**
   * Constructs a new instance.
   * 
   * {@see ExcelTableReader}
   */
  public StringExcelTableReader(String sheetName, Integer tableStartRowNumber,
      int tableStartColumnNumber, Integer tableRowSize, Integer tableColumnSize) {
    super(sheetName, tableStartRowNumber, tableStartColumnNumber, tableRowSize, tableColumnSize);
  }

  /**
   * Sets defaultDateTimeFormat.
   * 
   * @param dateTimeFormat dateTimeFormat string for {@link java.time.format.DateTimeFormatter}.
   * @return StringExcelTableReader (for method chain)
   */
  public StringExcelTableReader defaultDateTimeFormat(String dateTimeFormat) {
    getExcelReadUtil().setDefaultDateTimeFormat(dateTimeFormat);
    return this;
  }

  /**
   * Sets dateTimeFormat for specific column.
   * 
   * @param columnNumber the column number data is obtained from, 
   *     <b>starting with 1 and column A is equal to columnNumber 1</b>. 
   *     When the far left column of a table is 2 and you want to speciries the far left column,
   *     the columnNumber is 2.
   * @param dateTimeFormat dateTimeFormat string 
   *     for {@link java.time.format.DateTimeFormatter}.
   * @return StringExcelTableReader (for method chain)
   */
  public StringExcelTableReader columnDateTimeFormat(int columnNumber, String dateTimeFormat) {
    this.columnDateTimeFormatMap.put(columnNumber, dateTimeFormat);
    return this;
  }

  @Override
  public @Nullable String getDateTimeFormat(int columnNumber) {
    return columnDateTimeFormatMap.containsKey(columnNumber)
        ? columnDateTimeFormatMap.get(columnNumber)
        : null;
  }
}
