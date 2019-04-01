import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class XSSFExample1 {

    public static void main(String[] args) {
        String[] books = {"The Tempest", "Gitnjali", "Harry Potter"};
        String[] authors = {"William Shakespeare", "Rabindranath Tagore", "J. K. Rowling"};

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        sheet.setColumnWidth(0, (short)((50 * 80) / ((double) 1 / 20)));
        sheet.setColumnWidth(1, (short)((50 * 80) / ((double) 1 / 20)));
        workbook.setSheetName(0, "XSSFWorkbook example");

        Font font1 = workbook.createFont();
        font1.setFontHeightInPoints((short) 10);
        font1.setColor((short) 0xc);
        font1.setBold(true);
        XSSFCellStyle cellStyle1 = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle1.setFont(font1);

        Font font2 = workbook.createFont();
        font2.setFontHeightInPoints((short) 10);
        font2.setColor(Font.COLOR_NORMAL);
        XSSFCellStyle cellStyle2 = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle2.setFont(font2);

        Row headerRow = sheet.createRow(0);
        Cell cell1 = headerRow.createCell(0);
        cell1.setCellValue("Book");
        cell1.setCellStyle(cellStyle1);
        Cell cell2 = headerRow.createCell(1);
        cell2.setCellValue("Author");
        cell2.setCellStyle(cellStyle2);

        int rownum;
        Row row = null;
        Cell cell = null;

        for( rownum = 1; rownum < books.length; rownum++) {
            row = sheet.createRow(rownum);
            cell = row.createCell(0);
            cell.setCellValue(books[rownum-1]);
            cell.setCellStyle(cellStyle1);

            cell = row.createCell(1);
            cell.setCellValue(authors[rownum - 1]);
            cell.setCellStyle(cellStyle2);
        }

        try {
            final String FILE_NAME = "./xssf_example.xlsx";
            FileOutputStream fileOutputStream = new FileOutputStream(FILE_NAME);
            workbook.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            //workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
