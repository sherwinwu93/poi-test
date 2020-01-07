import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * Working with borders
 * @author wusd
 * @date 2020/1/6 14:29
 */
public class Borders {
    public static void main(String[] args) throws Exception {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue(4);
        // Style the cell with borders all around.
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());

        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());

        style.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());

        cell.setCellStyle(style);

        for (Sheet everysheet : wb ) {
            for (Row everyrow : everysheet) {
                for (Cell everycell : everyrow) {
                    // Do something here

                }
            }
        }

        try {
            OutputStream fileOut = new FileOutputStream("Borders.xls");
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
        wb.close();
    }
}
