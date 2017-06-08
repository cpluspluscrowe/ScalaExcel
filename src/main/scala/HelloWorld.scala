/**
  * Created by CCROWE on 6/8/2017.
  */
import java.io.FileOutputStream

import com.norbitltd.spoiwo.model.enums.CellFill
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.joda.time.LocalDate

object HelloWorld {

  def main(args: Array[String]) {

  }
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
  def CreateSheets(): Unit ={
    var wb = new XSSFWorkbook();
    var sheet1 = wb.createSheet("new sheet");
    var safeName:String = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]");
    var sheet3 = wb.createSheet(safeName);
    var fileOut = new FileOutputStream("workbook.xlsx")
    wb.write(fileOut);
    fileOut.close();
  }
}