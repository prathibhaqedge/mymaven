package Utilites;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtilities {
	Workbook wb;
	public ExcelUtilities(String excelpath)throws Throwable{
		FileInputStream fi = new FileInputStream(excelpath);
		wb= WorkbookFactory.create(fi); 
	}
	public int rowcount(String sheetname)
	{
		return wb.getSheet(sheetname).getLastRowNum();
	}
	public int cellcount(String sheetname)
	{
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}
	public String getcelldata(String sheetname,int row,int colums)
	{
		String data = " ";
		if(wb.getSheet(sheetname).getRow(row).getCell(colums).getCellType()==Cell.CELL_TYPE_NUMERIC) 
		{
			int celldata = (int) wb.getSheet(sheetname).getRow(row).getCell(colums).getNumericCellValue();
			data = String.valueOf(celldata);
			
		}
		else
		{
			data=wb.getSheet(sheetname).getRow(row).getCell(colums).getStringCellValue();
		}
		return data;
	}
	public void setcelldata(String sheetname,int row,int colums,String status,String writeexcelpath)throws Throwable
	{
		Sheet ws = wb.getSheet(sheetname);
		Row rown = ws.getRow(row);
		Cell celln=rown.createCell(colums);
		celln.setCellValue(status);
		if(status.equalsIgnoreCase("pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			font.setColor(IndexedColors.GREEN.getIndex());
			style.setFont(font);
			rown.getCell(colums).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			font.setColor(IndexedColors.RED.getIndex());
			style.setFont(font);
			rown.getCell(colums).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rown.getCell(colums).setCellStyle(style);
		}
		FileOutputStream fo = new FileOutputStream(writeexcelpath);
		wb.write(fo);
		}

}
