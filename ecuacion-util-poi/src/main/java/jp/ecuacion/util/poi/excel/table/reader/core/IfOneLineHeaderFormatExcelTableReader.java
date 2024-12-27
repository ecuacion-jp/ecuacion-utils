package jp.ecuacion.util.poi.excel.table.reader.core;

import jakarta.annotation.Nonnull;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.IfExcelTable;
import jp.ecuacion.util.poi.excel.table.IfOneLineHeaderFormatExcelTable;

/**
 * Is a reader which treats one line header format tables.
 * 
 * @param <T> See {@link IfExcelTable}.
 */
public interface IfOneLineHeaderFormatExcelTableReader<T>
    extends IfOneLineHeaderFormatExcelTable<T>, IfExcelTableReader<T> {

  @Override
  public default List<List<String>> updateAndGetHeaderList(@Nonnull List<List<T>> excelData) {
    List<String> list = excelData.remove(0).stream().map(el -> getStringValue(el)).toList();
    
    List<List<String>> rtnList = new ArrayList<>();
    rtnList.add(list);
    
    return rtnList;
  }
}
