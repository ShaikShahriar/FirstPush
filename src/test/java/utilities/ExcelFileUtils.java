package utilities;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.WebDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFileUtils
{
    Workbook wb;
    public ExcelFileUtils(String ExcelPath) throws IOException
    {
        FileInputStream fi= new FileInputStream(ExcelPath);
        wb= WorkbookFactory.create(fi);


    }

    //counting no of rows in a sheet
    public int rowCount(String sheetname)
    {
        return wb.getSheet(sheetname).getLastRowNum();

    }
   //get cell data from sheet
    public String getCellData(String sheetname,int row,int column) {
        String data = "";
        if (wb.getSheet(sheetname).getRow(row).getCell(column).getCellType() == CellType.NUMERIC) {
            int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
            data = String.valueOf(celldata);
        } else {
            data = wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
            4
        }
        return data;
    }
//method for writing
    public void setCellData(String sheetname,int row,int column,String status,String WriteExcel)throws Throwable
    {
        Sheet ws= wb.getSheet(sheetname);
        Row rowNum= ws.getRow(row);
        Cell cell= rowNum.createCell(column);
        cell.setCellValue(status);
        if (status.equalsIgnoreCase("pass"))
        {
            CellStyle style= wb.createCellStyle();
            Font font = wb.createFont();
            font.setColor(IndexedColors.GREEN.getIndex());
            font.setBold(true);
            style.setFont(font);
            rowNum.getCell(column).setCellStyle(style);
        }
        else if(status.equalsIgnoreCase("Fail"))
        {
            CellStyle style= wb.createCellStyle();
            Font font = wb.createFont();
            font.setColor(IndexedColors.RED.getIndex());
            font.setBold(true);
            style.setFont(font);
            rowNum.getCell(column).setCellStyle(style);

        } else if (status.equalsIgnoreCase("Blocked"))
        {
            CellStyle style= wb.createCellStyle();
            Font font = wb.createFont();
            font.setColor(IndexedColors.BLUE.getIndex());
            font.setBold(true);
            style.setFont(font);
            rowNum.getCell(column).setCellStyle(style);
        }
        FileOutputStream fo= new FileOutputStream(WriteExcel);
        wb.write(fo);
    }

}


