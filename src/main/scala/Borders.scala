import java.io.FileOutputStream

import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
  * Created by CCROWE on 6/8/2017.
  */
object Borders {
  def BordersChange(): Unit ={
    var wb = new XSSFWorkbook();
    var sheet = wb.createSheet("new sheet");
    var row = sheet.createRow(1);
    var cell = row.createCell(1);
    cell.setCellValue("Changing this cell's borders");

    var style = wb.createCellStyle();
    style.setBorderBottom(1:Short);
    style.setBorderLeft(1:Short);
    style.setBorderRight(1:Short);
    style.setBorderTop(1:Short);
    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
    style.setLeftBorderColor(IndexedColors.BLUE.getIndex());
    cell.setCellStyle(style);

    var fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }
}
