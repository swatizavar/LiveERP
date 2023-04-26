package utilities;

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

//used interface instead of class to make things simple and reuse code
public class ExcelFileUtil {
	Workbook wb;
	//Constructor for reading excel path
	public ExcelFileUtil(String ExcelPath) throws Throwable 
	{
		FileInputStream fi=new FileInputStream(ExcelPath);
		wb=WorkbookFactory.create(fi);
	}

	//Counting number of rows in sheet
	public int rowCount(String SheetName)
	{
		return wb.getSheet(SheetName).getLastRowNum();
	}

	//Read cell data
	public String getCellData(String SheetName,int row,int column)
	{
		String data="";
		if(wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)
		{
			//read integer type cell data
			int celldata=(int) wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue();
			data= String.valueOf(celldata);
		}
		else
		{
			data=wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}

	//Method for set cell data
	public void setCellData(String sheetName,int row,int column, String status,String WriteExcel) throws Throwable
	{
		//Get sheet from wb
		Sheet ws=wb.getSheet(sheetName);
		//get row from sheet
		Row rowNum=ws.getRow(row);
		//create cell in row
		Cell cell=rowNum.createCell(column);
		//write status
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("pass"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			//colour text green
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("fail"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			//colour text green
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(column).setCellStyle(style);
		}
		else if(status.equalsIgnoreCase("blocked"))
		{
			CellStyle style=wb.createCellStyle();
			Font font=wb.createFont();
			//colour text green
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			style.setFont(font);
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			rowNum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo=new FileOutputStream(WriteExcel);
		wb.write(fo);
	}
}
