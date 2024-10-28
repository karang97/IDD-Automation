package test_cases;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import groovy.lang.Newify;

public class Excel_copy_paste {
    public static void main(String[] args) throws IOException {
    	String sourceString="C:\\files\\SOURCE.xlsx";
    	FileInputStream inputStream=new FileInputStream(sourceString);
    	XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
    	XSSFSheet sheet=workbook.getSheetAt(0);
    	
    	Iterator iterator=  sheet.iterator();
    	
    	while(iterator.hasNext()) 
    	{
    		XSSFRow row=(XSSFRow) iterator.next();
    		
    		Iterator cellIterator=row.cellIterator();
    		
    		while(cellIterator.hasNext())
    		{
    			XSSFCell cell=(XSSFCell) cellIterator.next();
    			switch (cell.getCellType()) 
    			{
				case STRING: System.out.println(cell.getStringCellValue());break;
				case NUMERIC: System.out.println(cell.getNumericCellValue());break;
				case BOOLEAN: System.out.println(cell.getBooleanCellValue());break;
								
				}
    			System.out.println();
    		}
    		System.out.println();
    	}
    	
    	
    }
    }
  