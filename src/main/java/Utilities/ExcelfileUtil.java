package Utilities;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelfileUtil {
	Workbook wb;
	//constructor for reading excel path
	public ExcelfileUtil(String excelpath) throws Throwable {
		FileInputStream fi=new FileInputStream(excelpath);
		wb=WorkbookFactory.create(fi);

	}
	//row count in a sheet
	public int rowCount(String sheetname) {
		return wb.getSheet(sheetname).getLastRowNum();

	}
	//colume count in a row
	public int colCount(String sheetname) {
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}
	//get data from cell
	public String getCellData(String sheetname,int row,int column) {
		String data=null;
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC) {
			int celldata =(int)wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			//convert celldata into string type
			data=String.valueOf(celldata);
		}
		else {
			data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}
	//set cell data
	public void setCellData(String sheetname,int row,int column,String status,String writeexcel) throws Throwable {
		//get sheet from wb
		Sheet ws=wb.getSheet(sheetname);
		//get row from sheet
		Row rownum=ws.getRow(row);
		//creat cell in a row
		Cell cell=rownum.createCell(column);
		//set satatus in a cell
		cell.setCellValue(status);
		if(status.equalsIgnoreCase("Pass")) {
			//creat cell style
			CellStyle style=wb.createCellStyle();
			//creat front
			Font font=wb.createFont();
			//Apply color to the text
			font.setColor(IndexedColors.GREEN.getIndex());
			//Apply bold to the text
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);

		}
		else if(status.equalsIgnoreCase("Fail")){
			//creat cell style
			CellStyle style=wb.createCellStyle();
			//creat front
			Font font=wb.createFont();
			//Apply color to the text
			font.setColor(IndexedColors.RED.getIndex());
			//Apply bold to the text
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);

		}
		else if(status.equalsIgnoreCase("Not Executed")) {
			//creat cell style
			CellStyle style=wb.createCellStyle();
			//creat front
			Font font=wb.createFont();
			//Apply color to the text
			font.setColor(IndexedColors.BLUE.getIndex());
			//Apply bold to the text
			font.setBold(true);
			font.setBoldweight(font.BOLDWEIGHT_BOLD);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo=new FileOutputStream(writeexcel);
		wb.write(fo);
	}
}
