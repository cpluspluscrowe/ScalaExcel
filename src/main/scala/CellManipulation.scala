import java.io.FileOutputStream

import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
  * Created by CCROWE on 6/8/2017.
  */
object CellManipulation {
  def SetCellValues(): Unit ={
    var wb = new XSSFWorkbook();
    var createHelper = wb.getCreationHelper();
    var sheet = wb.createSheet("new sheet");
    var row = sheet.createRow(0);
    var cell = row.createCell(0);
    cell.setCellValue(1);
    row.createCell(1).setCellValue(1.2);
    row.createCell(2).setCellValue(
      createHelper.createRichTextString("This is a string")
    );
    row.createCell(3).setCellValue(true);
    var fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }
}
