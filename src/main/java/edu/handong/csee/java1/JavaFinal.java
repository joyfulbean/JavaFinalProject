package edu.handong.csee.java1;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.LinkedList;
import java.util.Queue;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.cli.Options;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.ParseException;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;
import com.opencsv.CSVWriter;

import edu.handong.csee.java2.Report1;
import edu.handong.csee.java2.Report2;
import edu.handong.csee.java3.DifferentColumException;
import edu.handong.csee.java3.GenQueue;

public class JavaFinal {
	String input;
	String output;
	boolean help;
	ArrayList<String> errorvalue = new ArrayList<String>();
	
	public void run(String[] args) throws IOException {
		
		//CLI Commons
		Options options = createOptions();
		
		if(parseOptions(options, args)){
			if (help){
				printHelp(options);
				return;
			}
		// path is required (necessary) data so no need to have a branch.
		System.out.println("You provided \"" + input + "\" as the value of the option p");
		
		//input the cli
		File a = new File(input);
		
		//store the list of zip file and file
		String[] nameofzipfile = new String[a.listFiles().length];
		//store the data
		//ArrayList<String> value = new ArrayList<String>();
		GenQueue<String> value;
		value = new GenQueue<String>();
		ArrayList<String> value2 = new ArrayList<String>();
		//Queue<String> errorvalue = new LinkedList<String>();
		
		//acess the zip file
		//list the all files and call the function to store the each lines by types to read
		int numdr = 0;
		int o = 0;
		for(File file: a.listFiles()) {
			 if(file.isFile()) {
				 nameofzipfile[o] = (file.getAbsolutePath());
				 System.out.println(nameofzipfile[o]);
				 
				 ZipFile zipFile = new ZipFile(nameofzipfile[o], "MS949");
				 
				 Enumeration<? extends ZipArchiveEntry> entries = zipFile.getEntries();
		          while(entries.hasMoreElements()){
		             ZipArchiveEntry entry = entries.nextElement();
		             if(entry.isDirectory()){
					    } 
		             else {
		            	 if(entry.getName().contains("(요약문)")) {
		            		 String zip = nameofzipfile[o];
		            		 InputStream stream = zipFile.getInputStream(entry);
		            		 for(String values:getData(stream,zip)) {
		            			 value.enqueue(values);
		            			 //value.add(values);
		            			 //System.out.println(values);
		            		 }
		            	 }
		            	else if(entry.getName().contains("(표.그림)")) {
		            		 String zip = nameofzipfile[o];
		            		InputStream stream = zipFile.getInputStream(entry);
		            		 for(String values:getData2(stream, zip)) {
		            			 value2.add(values);
		            			//System.out.println(values);
		            		 }
		            	}
		            	}
				     }
		          zipFile.close();
		          o++;
		        }  
			}
		
	String output1 = output+"output1.csv";
	String output2 = output+"output2.csv";
	String output3 = output+"error.csv";
	
	//System.out.println(output1);
		
			// print the output1.csv
			try (
					
				FileOutputStream outputStream = new FileOutputStream(output1);
			    BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(outputStream));

	            CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT
	                    .withHeader("ID", "Topic", "Summary", "Keyword","Date","Reference","Origin","CopyRight"));
	            
	            ){
				/*
				//System.out.println(output1);
				for (String strings : value) {
					csvPrinter.printRecord(strings);
					//System.out.println(strings);
				}*/
				{
					while (value.hasItems()) {
				         String strings = value.dequeue();
				         csvPrinter.printRecord(strings);
				         System.out.println(strings);
				      }
				}
	            	csvPrinter.flush();            	
	            }
			finally {

		      }
			// print the output2.csv
			
			try (
					FileOutputStream outputStream = new FileOutputStream(output2);
			       BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(outputStream));
	            //BufferedWriter writer = Files.newBufferedWriter(Paths.get(output1),Charset.forName("EUC-KR"));

	            CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT
	                    .withHeader("ID", "Title", "ID_Number", "Type","Explanation","Num"));
	            
	            ){
				for (String strings : value2) {
	            csvPrinter.printRecord(strings);
				}
	            csvPrinter.flush();            
	            }
			
			// print the error.csv
			try (
					FileOutputStream outputStream = new FileOutputStream(output3);
			       BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(outputStream));
	            //BufferedWriter writer = Files.newBufferedWriter(Paths.get(output1),Charset.forName("EUC-KR"));

	            CSVPrinter csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT
	                    .withHeader("Error_ZipFileName"));
	            
	            ){
				for (String strings : errorvalue) {
	            csvPrinter.printRecord(strings);
				}
	            csvPrinter.flush();            
	            }
			
			}
		}
				

		
		//인자로 번호 같이 전달해서, 앞에서붙이기?
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
			//System.out.println(line);
         	result.add(line);
			    }
		}
		catch(DifferentColumException e) {
			  System.out.println(e.getMessage());
			  errorvalue.add(nameofzipfile+"-type2");
			  //System.out.println(errorvalue);
		}
		return result;
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
		//System.out.println(line);
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

	

private boolean parseOptions(Options options, String[] args) {
	CommandLineParser parser = new DefaultParser();

	try {

		CommandLine cmd = parser.parse(options, args);

		input = cmd.getOptionValue("i");
		output = cmd.getOptionValue("o");
		help = cmd.hasOption("h");			

	} catch (Exception e) {
		printHelp(options);
		return false;
	}

	return true;
}



//Definition Stage
	private Options createOptions() {
		Options options = new Options();

		// add options by using OptionBuilder
		options.addOption(Option.builder("i").longOpt("input")
				.desc("Set an input file path")
				.hasArg()
				.argName("input path")
				.required()
				.build());

		// add options by using OptionBuilder
		options.addOption(Option.builder("o").longOpt("output")
				.desc("Set an output file path")
				.hasArg()     // this option is intended not to have an option value but just an option
				.argName("output path")
				.required() // this is an optional option. So disabled required().
				.build());
		
		// add options by using OptionBuilder
		options.addOption(Option.builder("h").longOpt("help")
		        .desc("Help")
		        .build());
		
		return options;
	}


private void printHelp(Options options) {
		// automatically generate the help statement
		HelpFormatter formatter = new HelpFormatter();
		String header = "JavaFinalProject";
		String footer = "" ; // Leave this empty.
		formatter.printHelp("JavaFinalProject", header, options, footer, true);
}

}






		
	
		
	
