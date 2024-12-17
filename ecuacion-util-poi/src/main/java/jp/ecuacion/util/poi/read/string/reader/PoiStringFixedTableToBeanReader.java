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
package jp.ecuacion.util.poi.read.string.reader;

import java.io.IOException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import jp.ecuacion.lib.core.exception.checked.AppException;
import jp.ecuacion.lib.core.util.BeanValidationUtil;
import jp.ecuacion.util.poi.enums.NoDataString;
import jp.ecuacion.util.poi.read.string.bean.PoiStringTableBean;
import org.apache.poi.EncryptedDocumentException;

/**
 * excel側のtableを、その構造を気にせず1つのbeanにそのまま格納する場合はこちらを使用可能。 readメソッドが用意されているためより便利。
 */
public abstract class PoiStringFixedTableToBeanReader<T extends PoiStringTableBean>
    extends PoiStringFixedTableReader {

  public PoiStringFixedTableToBeanReader() {
    super();
  }

  public PoiStringFixedTableToBeanReader(NoDataString noDataString) {
    super(noDataString);
  }

  public List<T> read(String filePath)
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
   * 本機能を使用した他appでのtestにおいて、readerのうちexcel読み込み部分のみを置き換えたい場合に本メソッドをoverrideする目的でメソッド分け。
   * 実際にexcelを読み込むテストは準備負荷が高いため、本メソッドを置き換えることで極力実excelファイルの読み込みを避けたテストを行うこと。
   */
  protected List<T> excelTableToBeanList(String filePath) throws AppException, IOException {
    List<List<String>> lines = getTableValues(filePath);

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

  /** 
   * 固定のテーブルなので、空行があっても読み込みを継続することは極めて考えにくいことから空行で終了とする。
   * 本メソッドをoverrideすることにより空行があっても読み込みを継続し固定の行数を読むことは可能。
   */
  @Override
  protected Integer getTableRowSize() {
    return null;
  }
}
