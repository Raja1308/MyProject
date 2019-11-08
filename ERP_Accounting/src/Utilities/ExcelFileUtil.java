package Utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
	Workbook wb;
//to read excel path


public ExcelFileUtil()throws Throwable
{
	FileInputStream fis=new FileInputStream("D:/Ojt4oclock/ERP_Accounting/TestInput/InputSheet.xlsx");
	wb=WorkbookFactory.create(fis);
}
//count no of rows in sheet
public int rowCount(String sheetname)
{
	return wb.getSheet(sheetname).getLastRowNum();
}
//count no of columns in row
public int colcount(String sheetname)
{
	return wb.getSheet(sheetname).getRow(0).getLastCellNum();
}
//reading data from cell
public String getCellData(String sheetname,int row,int column)
{
	String data="";
	if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC)	
	{
		int celldata=(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
		//convert celldata into String
		data=String.valueOf(celldata);
	}	
	else
	{
	data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();	
	}
	return data;
}
//write status into results column
public void setCellData(String sheetname,int row,int column,String Status)throws Throwable
{
	//get sheet
	Sheet ws=wb.getSheet(sheetname);
	//get row
	Row r=ws.getRow(row);
	//create column in row
	Cell cell=r.createCell(column);
	//write into status
	cell.setCellValue(Status);
	if(Status.equalsIgnoreCase("Pass"))
	{
		//create a cell style
		CellStyle style=wb.createCellStyle();
		//create a font
		Font font=wb.createFont();
		//apply color to the text
		font.setColor(IndexedColors.GREEN.getIndex());
		//apply bold to the text
		font.setBold(true);
		//set font
		style.setFont(font);
		//set cell style
		r.getCell(column).setCellStyle(style);
	}
	else if(Status.equalsIgnoreCase("Fail"))
	{
		//create a cell style
		CellStyle style=wb.createCellStyle();
		//create a font
		Font font=wb.createFont();
		//apply color to the text
		font.setColor(IndexedColors.RED.getIndex());
		//apply bold to the text
		font.setBold(true);
		//set font
		style.setFont(font);
		//set cell style
		r.getCell(column).setCellStyle(style);	
	}	
	else if(Status.equalsIgnoreCase("Not Exceuted"))
	{
		//create a cell style
		CellStyle style=wb.createCellStyle();
		//create a font
		Font font=wb.createFont();
		//apply color to the text
		font.setColor(IndexedColors.BLUE.getIndex());
		//apply bold to the text
		font.setBold(true);
		//set font
		style.setFont(font);
		//set cell style
		r.getCell(column).setCellStyle(style);	
	}
	FileOutputStream fos=new FileOutputStream("D:/Ojt4oclock/ERP_Accounting/TestOutput/Hybrid.xlsx");
	wb.write(fos);
	fos.close();
 }
}