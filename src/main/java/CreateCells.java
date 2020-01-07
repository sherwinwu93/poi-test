import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.Calendar;
import java.util.Date;

/**
 * @author wusd
 * @date 2020/1/6 11:06
 */
public class CreateCells {
    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();
//        Workbook wb = new XSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
        //Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(0);
        //Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        //Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                creationHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        //Write the output to a file
        try {
            OutputStream fileOut = new FileOutputStream("CreateCells.xls");
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void createDateCells() {
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        //Create a row and put some cells in it. Row are 0 based.
        Row row = sheet.createRow(0);

        //Create a cell and put a date value in it. The first cell is not styled
        //as a date.
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        //we style the second cell as a date(and time). It is important to
        //create a new cell style from the workbook otherwise you can end up
        //modifying the built in style and effecting not only this cell but other cells.
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                creationHelper.createDataFormat().getFormat("m/d/y h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        //you can also set date as java.util.Calendar
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        //Write the output to a file
        try {
            OutputStream fileOut = new FileOutputStream("CreateDateCells.xls");
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *
     */
    public static void differentTypeCell() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow(2);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        row.createCell(5).setCellType(CellType.BLANK);
        // Write the output to a file
        try {
            OutputStream fileOut = new FileOutputStream("workbook.xls");
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void useFileOrInputStream() {
        try {
            //Use a file
            Workbook wb = WorkbookFactory.create(new File("CreateCells.xls"));

            //Use an InputStream, needs more memory
            Workbook wb2 = WorkbookFactory.create(new FileInputStream("CreateCells.xls"));

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
