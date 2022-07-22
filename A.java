package app_java_10;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class A {

	 //excel reading using poi-api
	
		public static void main(String[] args) {
        	   ArrayList data = new ArrayList();
        	   
        		try {
			FileInputStream fis = new FileInputStream("D://PSA-JAVA TRAINING//EXCEL reading using POI-API//testread.xlsx");
			
			XSSFWorkbook Workbook = new XSSFWorkbook(fis);
		//	workbook.getSheet("sheet1");
			XSSFSheet sheetAt = Workbook.getSheetAt(0);
		  Iterator<Row> itr = sheetAt.iterator();
		  
		  while(itr.hasNext()) {
			  Row row = itr.next();
	  Iterator<Cell> cellIterator = row.cellIterator();
			 
			           while (cellIterator.hasNext()) {
			           Cell cell = cellIterator.next();			           
			          System.out.println(cell.getCellType());
			          
			       if(cell.getCellType()==CellType.STRING);{
			  //  System.out.println(cell.getStringCellValue());
			    	   data.add(cell.getStringCellValue());
			      
			         } else if (cell.getCellType()==CellType.NUMERIC); {
			        	   System.out.println(cell.getStringCellValue());
			        	   //typecast double value to int
			        	   double numericCellValue = cell.getNumericCellValue();
			         //      System.out.println((int)numericCellValue);
			              data.add((int)numericCellValue);
			           }
		  }
		  }
		  for(Object testData : data) {
			  if(testData.equals("pankaj"));
			  System.out.println(testData);
		  }
           }catch (Exception e) { 
				e.printStackTrace();
			}		

           
           }
           }
