package edu.handong.csee.java1;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Enumeration;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import edu.handong.csee.java2.Report1;
import edu.handong.csee.java2.Report2;
import edu.handong.csee.java3.DifferentColumException;
import edu.handong.csee.java3.GenQueue;

public class Zip extends Thread{
	
	GenQueue<String> value = new GenQueue<String>();
	GenQueue<String> value2 = new GenQueue<String>();
	ArrayList<String> errorvalue = new ArrayList<String>();
	
public GenQueue<String> readData(String nameofzipfile, ZipFile zipFile,ZipArchiveEntry entry) throws IOException{
       		 String zip = nameofzipfile;
       		 InputStream stream = zipFile.getInputStream(entry);
       		 for(String values:getData(stream,zip)) {
       			 value.enqueue(values);
       		 }
       		 return value;
}
	     

     
public GenQueue<String> readData2(String nameofzipfile, ZipFile zipFile,ZipArchiveEntry entry) throws IOException{
	
  		String zip = nameofzipfile;
  		InputStream stream = zipFile.getInputStream(entry);
  		 for(String values:getData2(stream, zip)) {
  			value2.enqueue(values);
  		 }
  		return value2;

}

     
     
 public ArrayList<String> getData(InputStream is, String nameofzipfile) throws IOException{
    		//define the field
    		ArrayList<String> result = new ArrayList<String>();
    		XSSFSheet sheet= null;
    		XSSFRow row = null;
    		XSSFCell cell = null;
    				
    		// read the workbook
    		XSSFWorkbook wb = new XSSFWorkbook(is);
    		sheet = wb.getSheetAt(0);

    		// number of rows in the sheet
    		int rows = ((XSSFSheet) sheet).getPhysicalNumberOfRows();
    		System.out.println(rows);
    		String line = "";
    		String temp = "";
    		Report1 report = new Report1();
    		try {	 
    		//for loops until the end of rows 
    		for (int r = 1; r < rows; r++) {
    			row = ((XSSFSheet) sheet).getRow(r);
    			if (row != null) { 
    				System.out.println(row.getLastCellNum());
    	            if(row.getLastCellNum() != 7) {
    	    			throw new DifferentColumException();
    	    		}
    	            int cells = row.getPhysicalNumberOfCells();
    	            for (int c = 0; c < cells; c++) {
    	                cell  = row.getCell(c);
    	                if(cell == null) {
    	                  temp = "";
    	                }
    	                else if(CellType.STRING == cell.getCellType()) {
    	                  temp = (((org.apache.poi.ss.usermodel.Cell) cell).getStringCellValue());
    	                }
    	                else if(CellType.NUMERIC == cell.getCellType()){
    	                  temp = Double.toString(((org.apache.poi.ss.usermodel.Cell) cell).getNumericCellValue());
    	                }
    	                if(c == 0){
    	                   report.setTopic(temp);    	
    	                }
    	                else if(c == 1) {
    	                	report.setSummary(temp);    	
    	                }
    	                else if(c == 2) {
    	                	report.setKeyword(temp);    	
    	                }
    	                else if(c == 3) {
    	                	report.setDate(temp);    	
    	                }
    	                else if(c == 4) {
    	                	report.setReference(temp);    	
    	                }
    	                else if(c == 5) {
    	                	report.setOrigin(temp);    	
    	                }
    	                else if(c == 6) {
    	                	report.setCopyright(temp);    	
    	                }        
    	             }
    	                 line = nameofzipfile+","+report.getTopic() +","+ report.getSummary()+","+report.getKeyword()+"," + report.getDate()+"," + report.getReference()+"," + report.getOrigin()+"," + report.getCopyright();
    			}
    	     	result.add(line);
    			    }
    		}
    				catch(DifferentColumException e) {
    					System.out.println(e.getMessage());
    					errorvalue.add(nameofzipfile+"-type1");
    					System.out.println(errorvalue);
    				}
    				return result;
    			}
 
 
 public ArrayList<String> getData2(InputStream is, String nameofzipfile) throws IOException{
		//define the field
		ArrayList<String> result = new ArrayList<String>();
		XSSFSheet sheet= null;
		XSSFRow row = null;
		XSSFCell cell = null;
				
		// read the workbook
		XSSFWorkbook wb = new XSSFWorkbook(is);
		sheet = wb.getSheetAt(0);

		// number of rows in the sheet
		int rows = ((XSSFSheet) sheet).getPhysicalNumberOfRows();
		try {	    
		String line = "";
		String temp = "";
		Report2 report = new Report2();
			
		//for loops until the end of rows 
		for (int r = 1; r < rows; r++) {
			row = ((XSSFSheet) sheet).getRow(r);
			if (row != null) { 
				//System.out.println(row.getLastCellNum());
	            if(row.getLastCellNum() != 5) {
	    			throw new DifferentColumException();
	    		}
	            int cells = row.getPhysicalNumberOfCells();
	            for (int c = 0; c < cells; c++) {
	                cell  = row.getCell(c);
	                if(cell == null) {
	                  temp = "";
	                }
	                else if(CellType.STRING == cell.getCellType()) {
	                  temp = (((org.apache.poi.ss.usermodel.Cell) cell).getStringCellValue());
	                }
	                else if(CellType.NUMERIC == cell.getCellType()){
	                  temp = Double.toString(((org.apache.poi.ss.usermodel.Cell) cell).getNumericCellValue());
	                }
	                if(c == 0){
	                   report.setTitle(temp);    	
	                }
	                else if(c == 1) {
	                	report.setIdnumber(temp);    	
	                }
	                else if(c == 2) {
	                	report.setType(temp);    	
	                }
	                else if(c == 3) {
	                	report.setExplanation(temp);    	
	                }
	                else if(c == 4) {
	                	report.setNum(temp);    	
	                }   
	             }
	                 line = nameofzipfile+","+report.getTitle() +","+ report.getIdnumber()+","+report.getType()+"," + report.getExplanation()+"," + report.getNum();
			}
      	result.add(line);
			    }
		}
		catch(DifferentColumException e) {
			  System.out.println(e.getMessage());
			  errorvalue.add(nameofzipfile+"-type2");
		}
		return result;
	}
		

    		
}
