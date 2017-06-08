import java.io.FileOutputStream
import java.util.{Calendar, Date}

import org.apache.poi.xssf.usermodel.{XSSFRow, XSSFWorkbook}

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
  def differentCellTypes(): Unit ={
    var wb:XSSFWorkbook = new XSSFWorkbook();
    var sheet = wb.createSheet("new sheet");
    var row:XSSFRow = sheet.createRow(2);
    row.createCell(0).setCellValue(1.1);
    row.createCell(1).setCellValue(new Date());
    row.createCell(2).setCellValue(Calendar.getInstance());
    row.createCell(3).setCellValue("a string");
    row.createCell(4).setCellValue(true);
    //row.createCell(5).setCellType(CellType.Error);
    var fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }
}
