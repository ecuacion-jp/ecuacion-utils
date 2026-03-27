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
package jp.ecuacion.util.pdfbox;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.pdfbox.excel.ExcelToPdf;
import org.apache.fontbox.util.BoundingBox;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.EncryptedDocumentException;

/**
 * 
 */
public class Main {

  /**
   * 
   */
  public static void main(String[] args) throws Exception {
    try (PDDocument doc = new PDDocument()) {
      PDPage page = new PDPage(PDRectangle.A4);
      doc.addPage(page);
      page.setRotation(90);
      page.getCropBox().
      int i = page.getRotation();
      page.setCropBox(new PDRectangle(new BoundingBox(0, 0, 100, 100)));
      try (PDPageContentStream cs = new PDPageContentStream(doc, page);) {
        PDFont font = PDType0Font.load(doc, new File("fonts/IPAexfont00401/ipaexm.ttf"));

        // y=0で文字列を出力すると、「g」など下側がある文字が切れる。
        // なので、下が切れないよう持ち上げておく
        float baseHeight = getPointFromMillimeter(10);
        String text = "ABCabcdefghijklpqrxyこんにちは、世界";
        float fontSize = 30;
        float textWidth = font.getStringWidth(text) * fontSize / 1000f;

        // 文字描画
        cs.beginText();
        cs.newLineAtOffset(0f, baseHeight);
        cs.setFont(font, fontSize);
        cs.showText(text);
        cs.endText();
      }

      doc.save("./helloworld.pdf");
    }
  }

  private static float getPointFromMillimeter(float millimeter) {
    return millimeter * 72f / 25.4f;
  }

  private void internalMain() throws EncryptedDocumentException, IOException, BizLogicAppException {
    new ExcelToPdf().execute("helloworld.pdf", "local-test/test01.xlsx", new String[] {"Sheet1"});
  }
}
