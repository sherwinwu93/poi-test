import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * How to create a new workbook
 *
 * @author wusd
 * @date 2020/1/6 10:58
 */
public class CreateSheet {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // You can use org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String newProposal)
        // for a safe way to create valid names.
        String safeName = WorkbookUtil.createSafeSheetName("[0'Brien's sales*?]");//returns "0'Brien's sales"
        Sheet sheet3 = wb.createSheet(safeName);
        try {
            OutputStream fileOut = new FileOutputStream("CreateSheet.xls");
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
