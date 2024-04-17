package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {


	Workbook wb;
	//constructor for reading excel path
	public ExcelFileUtil(String ExcelPath)throws Throwable
	{
		FileInputStream fi = new FileInputStream(ExcelPath);
		wb =WorkbookFactory.create(fi);
	}
	//count no of rows in a sheet
	public int rowCount(String sheetName)
	{
		return wb.getSheet(sheetName).getLastRowNum();
	}
	//method for reading cell data
	public String getCellData(String SheetName,int row,int column)
	{
		String data;
		if(wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
		{
			int celldata =(int)wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue();
			data =String.valueOf(celldata);
		}
		else
		{
			data = wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}
	//method for setcelldata
	public void setCellData(String sheetName,int row,int columns,String status,String WriteExcelpath)throws Throwable
	{
		//get sheet from wb
		Sheet ws =wb.getSheet(sheetName);
		//get row from sheet
		Row rowNum =ws.getRow(row);
		//create cell in row
		Cell cell =rowNum.createCell(columns);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			//colour with green
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Fail"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			//colour with red
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("Blocked"))
		{
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			//colour with blue
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			ws.getRow(row).getCell(columns).setCellStyle(style);
		}
		FileOutputStream fo =new FileOutputStream(WriteExcelpath);
		wb.write(fo);
	}
}