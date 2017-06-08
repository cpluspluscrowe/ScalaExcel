/**
  * Created by CCROWE on 6/8/2017.
  */
import java.io.{File, FileInputStream, FileOutputStream}
import java.util.{Calendar, Date}

import com.norbitltd.spoiwo.model.enums.CellFill
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFWorkbook}
import org.joda.time.LocalDate

object ScalaInExcel {
  def main(args: Array[String]) {
    var wb = WorkbookFactory.create(new File("workbook.xlsx"));
    wb = WorkbookFactory.create(new FileInputStream("workbook.xlsx"));
    var pkg = OPCPackage.open(new File("workbook.xlsx"))
    wb = new XSSFWorkbook(pkg);
    pkg.close();

  }
}