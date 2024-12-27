package jp.ecuacion.util.poi.excel.table;

import jakarta.annotation.Nullable;

/**
 * Provides the excel table with object type obtained from the excel data is {@code String}.
 */
public interface IfStringExcelTable  extends IfExcelTable<String> {

  @Override
  public default String getStringValue(@Nullable String cellData) {
    return cellData;
  }

}
