import java.io.FileOutputStream

import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
  * Created by CCROWE on 6/8/2017.
  */
object CellAlignment {
  def AlignCell(): Unit ={
    var wb = new XSSFWorkbook();
    var sheet =wb.createSheet();
    var row = sheet.createRow(2);
    row.setHeightInPoints(50);

    var cell = row.createCell(3);
    cell.setCellValue("Align It");
    var cellStyle = wb.createCellStyle();
    cellStyle.setAlignment(1:Short);
    cellStyle.setVerticalAlignment(1:Short);
    cell.setCellStyle(cellStyle);

    var fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }
}
