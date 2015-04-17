package com.cga.analyze.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
/**
 * analyze excel class
 * @author cga
 *
 */
public class AnalyzeExcel {
	//print logger 
	private static Logger logger = Logger.getLogger(AnalyzeExcel.class);
	
	/**
	 * 2003excel file analyze 
	 */
	@Test
	public void analyzeExcel(){
		String path = "C:/Users/hp/Desktop/11/2014秋学号.xls";
		File file = new File(path);
		Workbook book = null;
		try {
			//gain workbook object
			book = Workbook.getWorkbook(file);
			//gain sheets all
			Sheet[] sheets = book.getSheets();
			for (int i = 0; i < sheets.length; i++) {
				//gain rows all
				int rows = sheets[i].getRows();
				String[] titles = null;
				//gain title name
				titles = new String[sheets[i].getRow(0).length];
				//while rows all
				for (int j = 0; j < rows; j++) {
					//gain Cell[] all be rows[j]
					Cell[] cells = sheets[i].getRow(j);
					//declare titles[]
					for (int k = 0; k < cells.length; k++) {
						if(j == 0){
							String value = cells[k].getContents();
							titles[k] = value; 
						}else{
							String value = cells[k].getContents();
							System.out.print(titles[k] + ":" + value + " , ");
						}
					}
					System.out.println("   ");
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			logger.error(" analyze excel fail " + e.getMessage(),e);
		}finally{
			// close workbook objece
			if(book != null){
				book.close();
			}
		}
		
	}
	
	/**
	 * 2007Excel file analyze
	 */
	@Test
	public void analyzeExcel2007(){
		String path = "C:/Users/hp/Desktop/11/123.xlsx";
		File file = new File(path);
		XSSFWorkbook book = null;
		InputStream in = null;
		try {
			//gain  InputStream Object
			in = new FileInputStream(file);
			//gain 2007book object
			book = new XSSFWorkbook(in);
			//gain 2007sheets all
			int sheets = book.getNumberOfSheets();
			// while sheets numbers
			for (int i = 0; i < sheets; i++) {
				//gain one sheet
				XSSFSheet sheet = book.getSheetAt(i);
				//gain last row
				int rows = sheet.getLastRowNum();
				for (int j = 0; j < rows; j++) {
					//gain row
					XSSFRow row = sheet.getRow(j);
					//gain cells numbers
					int cells = row.getLastCellNum();
					for (int k = 0; k < cells; k++) {
						//gain cell value
						String value = row.getCell(k).getStringCellValue();
						System.out.print(value +  "   ");
					}
					System.out.println("");
				}
			}
		} catch (FileNotFoundException e) {
			logger.error("analyze excel2007 fail " + e.getMessage(),e);
		} catch (IOException e) {
			logger.error("analyze excel2007 fail " + e.getMessage(),e);
		}finally{
			if(in != null){
				try {
					in.close();
				} catch (IOException e) {
					logger.error("colse InputStream object fail " + e.getMessage(),e);
				}
			}
		}
		
	}
	
}
