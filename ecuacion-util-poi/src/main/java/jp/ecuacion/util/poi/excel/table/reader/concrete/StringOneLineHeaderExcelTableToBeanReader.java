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
package jp.ecuacion.util.poi.excel.table.reader.concrete;

import jakarta.annotation.Nonnull;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.annotation.RequireNonnull;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.PropertyFileUtil.Arg;
import jp.ecuacion.lib.core.util.ValidationUtil;
import jp.ecuacion.util.poi.excel.enums.NoDataString;
import jp.ecuacion.util.poi.excel.table.bean.StringExcelTableBean;
import org.apache.poi.EncryptedDocumentException;

/**
 * Stores the excel table data into a bean.
 */
public class StringOneLineHeaderExcelTableToBeanReader<T extends StringExcelTableBean>
    extends StringOneLineHeaderExcelTableReader {

  private Class<?> beanClass;

  /**
   * Constructs a new instance. the obtained value 
   *     from an empty cell is {@code null}.
   * 
   * <p>In most cases {@code null} is recommended 
   *     because {@code Jakarta Validation} annotations (like {@code Max}) 
   *     returns valid for {@code null}, but invalid for {@code ""}.</p>
   *     
   * @param beanClass the class of the generic parameter {@code T} is hard to obtain 
   *     especially the constructor of this class is called directly with setting a 
   *     class instead of T, like {@code List<Foo> list = 
   *     new StringOneLineHeaderExcelTableToBeanReader<Foo>(...)}.<br>
   *     See <a href="https://stackoverflow.com/questions/19860393/java-generics-obtaining-actual-type-of-generic-parameter">here</a>.
   */
  public StringOneLineHeaderExcelTableToBeanReader(Class<?> beanClass,
      @RequireNonnull String sheetName, @RequireNonnull String[] headerLabels,
      Integer tableStartRowNumber, int tableStartColumnNumber, Integer tableRowSize,
      @SuppressWarnings("unchecked") T... parameterClass) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize);
    this.beanClass = beanClass;
  }

  /**
   * Constructs a new instance with the obtained value from an empty cell.
   * 
   * @param beanClass the class of the generic parameter {@code T} is hard to obtain 
   *     especially the constructor of this class is called directly with setting a 
   *     class instead of T, like {@code List<Foo> list = 
   *     new StringOneLineHeaderExcelTableToBeanReader<Foo>(...)}.<br>
   *     See <a href="https://stackoverflow.com/questions/19860393/java-generics-obtaining-actual-type-of-generic-parameter">here</a>.
   * @param noDataString the obtained value from an empty cell. {@code null} or {@code ""}.
   */
  public StringOneLineHeaderExcelTableToBeanReader(Class<?> beanClass,
      @RequireNonnull String sheetName, @RequireNonnull String[] headerLabels,
      Integer tableStartRowNumber, int tableStartColumnNumber, Integer tableRowSize,
      @Nonnull NoDataString noDataString) {
    super(sheetName, headerLabels, tableStartRowNumber, tableStartColumnNumber, tableRowSize,
        noDataString);
    this.beanClass = beanClass;
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
    return readToBean(filePath, true);
  }

  /**
   * Obtains excel table in the form of {@code List<PoiStringTableBean>}
   *     and validate obtained values..
   * 
   * @param filePath excelPath
   * @param validates whether validation is enabled or not
   * @return the list of {@code PoiStringTableBean}.
   * @throws AppException AppException
   * @throws EncryptedDocumentException EncryptedDocumentException
   * @throws IOException IOException
   */
  public List<T> readToBean(String filePath, boolean validates)
      throws AppException, EncryptedDocumentException, IOException {
    final String msgId = "jp.ecuacion.util.poi.excel.reader.ValidationMessagePostfix.message";
    List<T> rtnList = excelTableToBeanList(filePath);

    if (validates) {
      for (T bean : rtnList) {
        // jakarta validation. excel data is usually not shown on displays,
        // so "setMessageWithItemName(true)" is used.
        ValidationUtil.validateThenThrow(bean,
            ValidationUtil.parameters().isMessageWithItemNames(true)
                .messagePostfix(Arg.message(msgId, Arg.strings(sheetName))));

        // data integrity check
        bean.afterReading();
      }
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

      try {
        @SuppressWarnings("unchecked")
        T bean = (T) beanClass.getConstructor(List.class).newInstance(line);

        rtnList.add(bean);

      } catch (Exception ex) {
        throw new RuntimeException(ex);
      }
    }
    return rtnList;
  }

  @SuppressWarnings("unchecked")
  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> defaultDateTimeFormat(
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.defaultDateTimeFormat(
        dateTimeFormat);
  }

  @SuppressWarnings("unchecked")
  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> columnDateTimeFormat(int columnNumber,
      DateTimeFormatter dateTimeFormat) {
    return (StringOneLineHeaderExcelTableToBeanReader<T>) super.columnDateTimeFormat(columnNumber,
        dateTimeFormat);
  }

  @Override
  public StringOneLineHeaderExcelTableToBeanReader<T> ignoresAdditionalColumnsOfHeaderData(
      boolean value) {
    this.ignoresAdditionalColumnsOfHeaderData = value;
    return this;
  }
}
