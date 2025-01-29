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
package jp.ecuacion.util.poi.sample.readstringfromcell;

import java.net.URL;
import java.nio.file.Path;
import java.time.format.DateTimeFormatter;
import java.util.List;
import jp.ecuacion.util.poi.excel.table.reader.concrete.StringOneLineHeaderExcelTableToBeanReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ReadStringsFromOneLineHeaderFormatExcelTableToBeanSample {

  private static final String[] headerLabels =
      new String[] {"ID", "name", "date of birth", "age", "nationality"};

  private static final int HEADER_START_ROW = 3;
  private static final int START_COL = 2;

  public static void main(String[] args) throws Exception {

    Logger logger =
        LoggerFactory.getLogger(ReadStringsFromOneLineHeaderFormatExcelTableToBeanSample.class);

    logger.info("Procedure started.");

    // read
    List<SampleTableBean> beanList = read();

    beanList.stream().forEach(bean -> logger.info(bean.toString()));

    logger.info("Procedure finshed.");
  }

  private static List<SampleTableBean> read() throws Exception {

    // Get the path of the excel file.
    URL sourceUrl = ReadStringsFromOneLineHeaderFormatExcelTableToBeanSample.class.getClassLoader()
        .getResource("sample.xlsx");
    Path sourcePath = Path.of(sourceUrl.toURI()).toAbsolutePath();

    // Get the table data.
    return new StringOneLineHeaderExcelTableToBeanReader<SampleTableBean>(SampleTableBean.class,
        "Member", headerLabels, HEADER_START_ROW, START_COL, null)
            .defaultDateTimeFormat(DateTimeFormatter.ofPattern("MM/dd/yyyy"))
            .readToBean(sourcePath.toString());
  }
}
