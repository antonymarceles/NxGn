package com.fragrancial.dbUpdate;

import java.io.File;
import java.security.Timestamp;
import java.sql.Date;
import java.util.Calendar;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import com.excel.classes.excelreader;


public class fragrancialClean {

	public static void main(String[] args) throws Exception {
		
		dbConnection db = new dbConnection();
		db.connectDB();
		
		try {
	      String sql_postmeta = "DELETE FROM `wp_posts` where id >100"; 
		  db.executeQuery(sql_postmeta);
		  
	      String sql_postmeta1 = "DELETE FROM `wp_postmeta` where meta_id >100"; 
		  db.executeQuery(sql_postmeta1);
			  
	    } catch (BiffException e) {
	    	e.printStackTrace();
	    }		
	    
		System.out.println("Cleaned Database");
		db.closeDB();
	}
}
