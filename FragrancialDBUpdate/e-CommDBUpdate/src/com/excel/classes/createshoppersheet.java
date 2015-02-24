package com.excel.classes;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class createshoppersheet {
    
	  public static void main(String[] args) throws WriteException, IOException {
		  
		  	//Handle Input File
		  	String inputFile = "C:/personal/shopper_in.xls";
			excelreader excelReader = new excelreader();
			excelReader.setInputFile(inputFile);
			
			//Handle Output File 
		  	String outputFile = "C:/personal/shopper_in - Copy.xls";
		    excelwriter excelWriter = new excelwriter();
		    excelWriter.setOutputFile(outputFile);
		    
		    
		    
		    //Prepare Output File
		    File outputFileLoc = new File(outputFile);
		    WorkbookSettings wbSettings = new WorkbookSettings();
		
		    wbSettings.setLocale(new Locale("en", "EN"));
		
		    WritableWorkbook workbook = Workbook.createWorkbook(outputFileLoc, wbSettings);
		    workbook.createSheet("productData", 0);
		    WritableSheet excelSheet = workbook.getSheet(0);
		    
		    // Create Headers for New Output File
		    excelWriter.createLabel(excelSheet);
		    
		    
		    //Open & Read Input File
		    File inputWorkbook = new File(inputFile);
		    Workbook w;
		    try {
			      w = Workbook.getWorkbook(inputWorkbook);
			      // Get the first sheet
			      Sheet sheet = w.getSheet(0);
			     
			      // Loop over Columns & Rows and Write Data to Output File
			      for (int j = 0; j < sheet.getColumns(); j++) {
			    	  for (int i = 1; i < sheet.getRows(); i++) {
			    		  Cell cell = sheet.getCell(j, i);
			    		  excelWriter.addLabel(excelSheet, j, i, cell.getContents());
			    	  }
			      }
			      
		    } catch (BiffException e) {
		    	e.printStackTrace();
		    }
		    
		    workbook.write();
		    workbook.close();		    
	  }
    
}
