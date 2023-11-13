package phucho.AutoClickExcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelHelpers {
    private FileInputStream fis;
    private FileOutputStream fileOut;
    private Workbook wb;
    private Sheet sh;
    private Cell cell;
    private Row row;
    private CellStyle cellstyle;
    private Color mycolor;
    private String excelFilePath;
    private Map<String, Integer> columns = new HashMap<>();

    public void setExcelFile(String ExcelPath, String SheetName) throws Exception {
        try {
            File f = new File(ExcelPath);

            if (!f.exists()) {
                f.createNewFile();
                System.out.println("File doesn't exist, so created!");
            }

            fis = new FileInputStream(ExcelPath);
            wb = WorkbookFactory.create(fis);
            sh = wb.getSheet(SheetName);
            //sh = wb.getSheetAt(0); //0 - index of 1st sheet
            if (sh == null) {
                sh = wb.createSheet(SheetName);
            }

            this.excelFilePath = ExcelPath;

            //adding all the column header names to the map 'columns'
            sh.getRow(0).forEach(cell -> {
                columns.put(cell.getStringCellValue(), cell.getColumnIndex());
            });

        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    //set by column index
    public void setCellData(String text, int rowNumber, int colNumber) {
        try {
            row = sh.getRow(rowNumber);
            if (row == null) {
                row = sh.createRow(rowNumber);
            }
            cell = row.getCell(colNumber);

            if (cell == null) {
                cell = row.createCell(colNumber);
            }
            cell.setCellValue(text);

            XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
            style.setFillPattern(FillPatternType.NO_FILL);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);

            cell.setCellStyle(style);

            fileOut = new FileOutputStream(excelFilePath);
            wb.write(fileOut);
            fileOut.flush();
            fileOut.close();
        } catch (Exception e) {
            e.getMessage();
        }
    }
}
