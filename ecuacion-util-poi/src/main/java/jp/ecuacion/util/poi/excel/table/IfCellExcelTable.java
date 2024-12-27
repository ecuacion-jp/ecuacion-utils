package jp.ecuacion.util.poi.excel.table;

import jakarta.annotation.Nullable;
import jp.ecuacion.util.poi.excel.util.ExcelReadUtil;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Provides the excel table with object type obtained from the excel data is {@code Cell}.
 */
public interface IfCellExcelTable extends IfExcelTable<Cell> {

  @Override
  public default String getStringValue(@Nullable Cell cellData) {
    return cellData == null ? null : new ExcelReadUtil().getStringFromCell(cellData);
  }

}
