package excelProject;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
public class ExcelTrial
{
public static void main(String args[])

	{        
	
		String filePath ="f:\\dbreport.xls"; 
		FileOutputStream fileOut = null;

		try {
		        fileOut = new FileOutputStream(filePath);
		} catch (FileNotFoundException e1) {
			
			e1.printStackTrace();
		} 
		
	
		HSSFWorkbook wb = new HSSFWorkbook();
		 Sheet sheet = wb.createSheet();
		 
		  Row r1 = sheet.createRow(0);
		  Cell c1 = r1.createCell(0);
		  c1.setCellValue("student Name");
		  Cell c2 = r1.createCell(1);
		  c2.setCellValue("Roll no.");
		   
		  Row r2 = sheet.createRow(1);
		  Cell c3 = r2.createCell(0);
		  c3.setCellValue("Ravi");
		  Cell c4 = r2.createCell(1);
		  c4.setCellValue("564");
		    
		  Row r3 = sheet.createRow(2);
		  Cell c5 = r3.createCell(0);
		  c5.setCellValue("RAM");
		  Cell c6 = r3.createCell(1);
		  c6.setCellValue("783"); 
		  try { 
		        
			    //hello
		         wb.write(fileOut);
		         fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		    
		
		
	}//End of doGet

}
