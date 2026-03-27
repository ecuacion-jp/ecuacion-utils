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
package jp.ecuacion.util.pdfbox.excel.internal;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Gives a coordinat-points in two coordinates.
 * 
 * <p>When you create PDF (0, 0) is at the bottom-left point of the page.
 *     The point goes towards the right as the value of x becomes greater
 *     and the point goes towards the top as the value of y becomes greater.</p>
 *
 * <p>On the other hand, when you create a document, you start with top-left of it.
 *     The point wants to go towards the right as the value of x becomes greater
 *     and the point wants to go towards the bottom as the value of y becomes greater.<br>
 *     So coordinates of PDF is not very useful to us.</p>
 * 
 * <p>To resolve that situation we introduce this class 
 *     to translate document-standard coorrdinates to PDF coordinates.</p>
 *     
 * <p>
 */
public class CoordinatesManager {
  
  private Workbook workbook;
  private Sheet sheet;
  
  public CoordinatesManager(Workbook workbook, Sheet sheet) {
    String printArea = workbook.getPrintArea(workbook.getSheetIndex(sheet));
    short scale = sheet.getPrintSetup().getScale();
    boolean isLandscape = sheet.getPrintSetup().getLandscape();
    double headerMargin = sheet.getPrintSetup().getHeaderMargin();
    double footerMargin = sheet.getPrintSetup().getFooterMargin();
    double leftMargin = sheet.getMargin(PageMargin.LEFT);
    double rightMargin = sheet.getMargin(PageMargin.RIGHT);
    double topMargin = sheet.getMargin(PageMargin.TOP);
    double bottomMargin = sheet.getMargin(PageMargin.BOTTOM);
    int paperSize = sheet.getPrintSetup().getPaperSize();
  }
  
  public float getPdfCoordinatesYaxisValue(float excelCoordinatesYaxisValue) {
    return 1;
  }
}
