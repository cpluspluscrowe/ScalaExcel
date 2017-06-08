import java.io.FileOutputStream
import java.util.{Calendar, Date}

import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFWorkbook}

/**
  * Created by CCROWE on 6/8/2017.
  */
object Dates {
  def AddDates(): Unit ={
    var wb = new XSSFWorkbook();
    var createHelper = wb.getCreationHelper();
    var sheet = wb.createSheet("new sheet");
    var row:XSSFRow = sheet.createRow(0);
    var cell:XSSFCell = row.createCell(0);
    cell.setCellValue(new Date());

    var cellStyle = wb.createCellStyle();
    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
    cell = row.createCell(1);
    cell.setCellValue(new Date());
    cell.setCellStyle(cellStyle);

    cell = row.createCell(2);
    cell.setCellValue(Calendar.getInstance());
    cell.setCellStyle(cellStyle);

    var fileOut = new FileOutputStream("workbook.xlsx");
    wb.write(fileOut);
    fileOut.close();
  }
}
