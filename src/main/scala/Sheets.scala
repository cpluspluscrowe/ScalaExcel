import java.io.FileOutputStream

import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
  * Created by CCROWE on 6/8/2017.
  */
object Sheets {
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
