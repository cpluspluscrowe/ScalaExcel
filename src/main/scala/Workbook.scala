import java.io.{File, FileInputStream, FileOutputStream}

import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
  * Created by CCROWE on 6/8/2017.
  */
object Workbook {
  def FileWorkbook(): Unit ={ //full control over workbook lifecycle
    var wb = WorkbookFactory.create(new File("workbook.xlsx"));
    wb = WorkbookFactory.create(new FileInputStream("workbook.xlsx"));
    var pkg = OPCPackage.open(new File("workbook.xlsx"))
    wb = new XSSFWorkbook(pkg);
    pkg.close();
  }
}
