import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * @author wusd
 * @date 2020/1/6 14:10
 */
public class AlignmentOptions {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook(); //or new HSSFWorkbook();

        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(2);
        row.setHeightInPoints(30);

        createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
        createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
        createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

        try {
            OutputStream fileOut = new FileOutputStream("AlignmentOptions.xls");
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb     the workbook
     * @param row    the row to create the cell in
     * @param column the column number to create the cell in
     * @param halign the horizontal alignment for the cell.
     * @param valign valign the vertical alignment for the cell.
     */
    private static void createCell(Workbook wb, Row row, int column,
                                   HorizontalAlignment halign, VerticalAlignment valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);

        cell.setCellStyle(cellStyle);
    }
}
