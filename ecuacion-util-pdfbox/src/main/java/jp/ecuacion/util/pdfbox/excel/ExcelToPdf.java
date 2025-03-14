package jp.ecuacion.util.pdfbox.excel;

import java.io.File;
import java.io.IOException;
import jp.ecuacion.lib.core.exception.checked.BizLogicAppException;
import jp.ecuacion.util.pdfbox.excel.internal.CoordinatesManager;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.font.Standard14Fonts;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Creates a PDF file from an excel file.
 */
public class ExcelToPdf {

  private boolean showsMarginAreas;

  public void setShowsMarginAreas(boolean showsMarginAreas) {
    this.showsMarginAreas = showsMarginAreas;
  }

  /**
  * Creates a PDF file from an excel file.
   * 
   * @param excelPath input excel file path
   * @param sheetNames excel sheet names to print
   * @throws IOException IOException
   * @throws BizLogicAppException BizLogicAppException
   */
  public void execute(String excelPath, String[] sheetNames)
      throws IOException, BizLogicAppException {
    String excelName = new File(excelPath).getName();
    String pdfPath = new File(excelPath).getParent()
        .concat(excelName.substring(0, excelName.lastIndexOf(".")) + ".pdf");

    execute(pdfPath, excelPath, sheetNames);
  }

  /**
  * Creates a PDF file from an excel file.
   * 
   * @param pdfPath created PDF file path
   * @param excelPath input excel file path
   * @param sheetNames excel sheet names to print
   * @throws IOException IOException
   * @throws BizLogicAppException BizLogicAppException
   */
  public void execute(String pdfPath, String excelPath, String[] sheetNames)
      throws IOException, BizLogicAppException {
    PrintTarget[] printTargets = new PrintTarget[sheetNames.length];
    for (int i = 0; i < sheetNames.length; i++) {
      printTargets[i] = new PrintTarget(sheetNames[i]);
    }

    execute(pdfPath, excelPath, printTargets);
  }

  /**
  * Creates a PDF file from an excel file.
   * 
   * @param pdfPath created PDF file path
   * @param excelPath input excel file path
   * @param printTargets printTargets to print
   * @throws IOException IOException
   * @throws BizLogicAppException BizLogicAppException
   */
  public void execute(String pdfPath, String excelPath, PrintTarget[] printTargets)
      throws IOException, BizLogicAppException {

    Workbook wb = WorkbookFactory.create(new File(excelPath), null, true);

    PDDocument document = new PDDocument();

    for (PrintTarget printTarget : printTargets) {
      Sheet sheet = wb.getSheet(printTarget.getSheetName());
      if (sheet == null) {
        throw new BizLogicAppException("jp.ecuacion.util.pdfbox.excel.SheetNotExist.message");
      }

      CoordinatesManager manager = new CoordinatesManager(wb, sheet);

      int i0 = sheet.getColumnWidth(0);
      int i1 = sheet.getColumnWidth(1);
      int i2 = sheet.getColumnWidth(2);
      int i3 = sheet.getColumnWidth(3);
      int i4 = sheet.getColumnWidth(4);
      int i5 = sheet.getColumnWidth(5);

      PDPage page = new PDPage(PDRectangle.A5);
      document.addPage(page);

      if (showsMarginAreas) {
        showMarginAreas(wb, sheet, document, page, manager);
      }

      float f = page.getMediaBox().getWidth();
      float h = page.getMediaBox().getHeight();

      PDFont font = new PDType1Font(Standard14Fonts.FontName.COURIER);

      PDPageContentStream contentStream = new PDPageContentStream(document, page);
      contentStream.addRect(100f, 400f, 100f, 150f);
      contentStream.setNonStrokingColor(0.8F, 0.9F, 1F);
      contentStream.fill();
      contentStream.close();

      // contentStream.beginText();
      // contentStream.setFont(font, 12);
      // contentStream.newLineAtOffset(0f, 0f);
      // contentStream.showText("Hello World");
      // contentStream.endText();

      document.save(pdfPath);

      System.out.println("");
    }

    document.close();
  }

  private void showMarginAreas(Workbook workbook, Sheet sheet, PDDocument document, PDPage page,
      CoordinatesManager manager) throws IOException {

    PDPageContentStream contentStream = new PDPageContentStream(document, page);

    Pair<Short, Short> size = getVerticalAndHorizontalSize(sheet);
    short paperSize = sheet.getPrintSetup().getPaperSize();
    sheet.getPrintSetup().getLandscape();

    float f = page.getMediaBox().getWidth();
    float h = page.getMediaBox().getHeight();

    contentStream.addRect(100f, 400f, 100f, 150f);
    contentStream.setNonStrokingColor(0.8F, 0.9F, 1F);
    contentStream.fill();
    contentStream.close();
  }

  private Pair<Short, Short> getVerticalAndHorizontalSize(Sheet sheet) {
    return null;
  }

  /**
   * Stores the sheet and its pages to print.
   */
  public static class PrintTarget {
    private String sheetName;

    /**
     * 
     */
    int[] pageNumbers;

    /**
     * Constructs a new instance.
     * 
     * @param sheetName sheetName
     */
    public PrintTarget(String sheetName) {
      this.sheetName = sheetName;
      this.pageNumbers = new int[] {};
    }

    /**
     * Constructs a new instance.
     * 
     * @param sheetName sheetName
     * @param pageNumbers pageNumbers
     */
    public PrintTarget(String sheetName, int[] pageNumbers) {
      this.sheetName = sheetName;
      this.pageNumbers = pageNumbers;
    }

    public String getSheetName() {
      return sheetName;
    }

    public int[] getPageNumbers() {
      return pageNumbers;
    }
  }
}
