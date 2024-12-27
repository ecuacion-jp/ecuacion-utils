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
package jp.ecuacion.util.poi.excel.table.reader.string;

import jakarta.annotation.Nonnull;
import java.io.IOException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.BeanValidationUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.reader.string.bean.StringTableBean;
import org.apache.poi.EncryptedDocumentException;

/**
 * Stores the excel table data into a bean.
 */
public class StringOneLineHeaderFormatExcelTableToBeanReader<T extends StringTableBean>
    extends StringOneLineHeaderFormatExcelTableReader {

  /**
   * Constructs a new instance. the obtained value 
   *     from an empty cell is {@code null}.
   * 
   * <p>In most cases {@code null} is recommended 
   *     because {@code Bean Validation} annotations (like {@code Max}) 
   *     returns valid for {@code null}, but invalid for {@code ""}.</p>
   */
  public StringOneLineHeaderFormatExcelTableToBeanReader(@RequireNonnull String sheetName,
      @Nonnull String[] headerLabels, Integer tableStartRowNumber, int tableStartColumnNumber,
      Integer tableRowSize) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize);
  }

  /**
   * Constructs a new instance with the obtained value from an empty cell.
   * 
   * @param noDataString the obtained value from an empty cell. {@code null} or {@code ""}.
   */
  public StringOneLineHeaderFormatExcelTableToBeanReader(@RequireNonnull String sheetName,
      @Nonnull String[] headerLabels, Integer tableStartRowNumber, int tableStartColumnNumber,
      Integer tableRowSize, @Nonnull NoDataString noDataString) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        noDataString);
  }

  /**
   * Obtains excel table in the form of {@code List<PoiStringTableBean>}
   *     and validate obtained values..
   * 
   * @param filePath excelPath
   * @return the list of {@code PoiStringTableBean}.
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public List<T> readToBean(String filePath)
      throws AppException, EncryptedDocumentException, IOException {
    List<T> rtnList = excelTableToBeanList(filePath);

    // data check
    BeanValidationUtil valUtil = new BeanValidationUtil();
    for (T bean : rtnList) {
      // bean validation
      valUtil.validateThenThrow(bean);

      // dat整合性check
      bean.dataConsistencyCheck();
    }

    return rtnList;
  }

  /** 
   * Obtains excel table in the form of {@code List<PoiStringTableBean>}.
   * 
   * <p>Since {@code read} method contains validation function, 
   * you may want to test the validations.
   * Although you don't want to prepare excel files to test each case.<br>
   * When that's the case, you can skip the preparation of excel files 
   * by overriding this method and return list you want to test.</p>
   */
  protected List<T> excelTableToBeanList(String filePath) throws AppException, IOException {
    List<List<String>> lines = read(filePath);

    List<T> rtnList = new ArrayList<>();
    for (List<String> line : lines) {

      // new T() したいので、それを実現するためreflectionをこねくり回す
      // https://nagise.hatenablog.jp/entry/20131121/1385046248
      try {
        // 実行時の型が取れる。ここではHogeDaoなど
        Class<?> clazz = this.getClass();
        // ここではBaseDao<Hoge>がとれる
        Type type = clazz.getGenericSuperclass();
        ParameterizedType pt = (ParameterizedType) type;
        // BaseDaoの型変数に対するバインドされた型がとれる
        Type[] actualTypeArguments = pt.getActualTypeArguments();
        @SuppressWarnings("unchecked")
        Class<T> entityClass = (Class<T>) actualTypeArguments[0];
        T bean = (T) entityClass.getConstructor(List.class).newInstance(line);

        rtnList.add(bean);

      } catch (Exception ex) {
        throw new RuntimeException(ex);
      }
    }
    return rtnList;
  }
}
