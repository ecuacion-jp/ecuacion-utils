package jp.ecuacion.util.poi.excel.table.reader.concrete;

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

  protected String defaultDateFormat;
  protected Map<Integer, String> columnDateFormatMap = new HashMap<>();

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
   * Sets defaultDateFormat.
   * 
   * @param dateFormat dateFormat string for {@link java.text.SimpleDateFormat}.
   * @return ReturnUrlBean (for method chain)
   */
  public StringExcelTableReader defaultDateFormat(String dateFormat) {
    this.defaultDateFormat = dateFormat;
    return this;
  }

  /**
   * Sets dateFormat for specific column.
   * 
   * @param columnNumber the column number data is obtained from, 
   *     <b>starting with 1 and column A is equal to columnNumber 1</b>. 
   *     When the far left column of a table is 2 and you want to speciries the far left column,
   *     the columnNumber is 2.
   * @param dateFormat dateFormat string for {@link java.text.SimpleDateFormat}.
   * @return ReturnUrlBean (for method chain)
   */
  public StringExcelTableReader columnDateFormat(int columnNumber, String dateFormat) {
    this.columnDateFormatMap.put(columnNumber, dateFormat);
    return this;
  }

  @Override
  public String getDateFormat(int columnNumber) {
    return columnDateFormatMap.containsKey(columnNumber) ? columnDateFormatMap.get(columnNumber)
        : defaultDateFormat;
  }
}
