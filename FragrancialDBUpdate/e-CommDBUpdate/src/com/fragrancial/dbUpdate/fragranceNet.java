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

public class fragranceNet {

	public static void main(String[] args) throws Exception {
		
		dbConnection db = new dbConnection();
		db.connectDB();
		
	  	//Handle Input File
	  	String inputFile = "datafeed2.xls";
		excelreader excelReader = new excelreader();
		excelReader.setInputFile(inputFile);
				
	    //Open & Read Input File
	    File inputWorkbook = new File(inputFile);
	    Workbook w;
	    String sql_postmeta ;
	    
	    //Reset Stock Count to 0 for all products
	    String sql_resetData = "UPDATE wp_postmeta SET meta_value = '0' WHERE meta_key = '_stock' and post_id in  (select * from (SELECT post_id FROM wp_postmeta wp_pm where meta_value like '%fragranceNet%' and meta_key = 'image_link1') mockalias )"; 
        db.executeQuery(sql_resetData);

	    //Receive Date/Time fields for Database
		Calendar calendar = Calendar.getInstance();
		java.sql.Timestamp ourJavaTimestampObject = new java.sql.Timestamp(calendar.getTime().getTime());
		String datetimestamp = String.valueOf(ourJavaTimestampObject).substring(0, 19);
		    
		    
	    try {
		      w = Workbook.getWorkbook(inputWorkbook);
		      
		      // Get the first sheet
		      Sheet sheet = w.getSheet(0);		
		      String post_content = "", post_title = "", post_excerpt = "", post_product_name="", post_product_type="", post_category="", post_name = "", post_brand="", post_sku = "", post_image="", post_regularprice="", post_price=" ", post_saleprice="", post_stock="";
		      
		      // Loop over Columns & Rows and Write Data to Output File
	    	  for (int i = 1; i <= sheet.getRows(); i++) {
	    		    
	    		  String sql_posts = "INSERT INTO wp_posts (post_author, post_date, post_date_gmt, post_content, post_title, post_excerpt, post_status, comment_status, ping_status, post_password, post_name, to_ping, pinged, post_modified, post_modified_gmt, post_content_filtered, post_parent, guid, menu_order, post_type, post_mime_type, comment_count) Values ";

    			  if (!(sheet.getCell(1, i).getContents().equals("Fragrance") && (sheet.getCell(2, i).getContents().equals("Fragrances") || sheet.getCell(2, i).getContents().equals("Gift Sets")))) {
    				  if (!(sheet.getCell(1, i).getContents().equals("Aromatherapy") || (sheet.getCell(1, i).getContents().equals("Candle")))) {
    			
								post_sku = sheet.getCell(0, i).getContents();
			    			    
								post_brand = sheet.getCell(5, i).getContents().replace("'", "");
								post_title = sheet.getCell(3, i).getContents().replace("'", "") + " - " + sheet.getCell(6, i).getContents().replace("'", "");
								
								if (sheet.getCell(8, i).getContents().equals("") || sheet.getCell(8, i).getContents().equals(null)) {
									post_content = toCamelCase(post_title);	
									post_excerpt = "";
								} else {
									post_content = "<b>" + toCamelCase(sheet.getCell(3, i).getContents().replace("'", "")) + "</b>: " + sheet.getCell(8, i).getContents().replace("'", "");
									post_excerpt = sheet.getCell(8, i).getContents().replace("'", "").split("\\.")[0].concat(".");			
								}
								
								post_name = (sheet.getCell(3, i).getContents().replace("'", "")+" - "+sheet.getCell(0, i).getContents()).replace("'", "").replace(" - ","-").replace(" -- ","-").replace(" + ", "-").replace("?", "-").toLowerCase().replace(" .", "-").replace(".", "-").replace(" / ", "-").replace("/","-").replace("(", "").replace(")","").replace("&", "-").replace("%", "-").replace("+", "-").replace("!", "-");
								post_name = (post_name).replace(" ", "-").replace("----","-").replace("---","-").replace("--","-");

								post_product_name = sheet.getCell(3, i).getContents().replace("'", "");
								post_product_type = sheet.getCell(6, i).getContents().replace("'", "");
			
								//Setting & Manipulating Prices
								post_regularprice = sheet.getCell(11, i).getContents();
								
								Double convertorVal = ((Double.parseDouble(sheet.getCell(12, i).getContents()))*(1.28+(Math.random()* (.04)))) * 100;						
								post_saleprice = String.valueOf(Math.ceil(convertorVal/100)-.01);
											
								if (post_regularprice.equals("") || post_regularprice.equals(null)) {
									post_regularprice = "-";
								} else if (Double.parseDouble(post_regularprice) < Double.parseDouble(post_saleprice)) {
									post_saleprice = post_regularprice;
									post_regularprice = "-";  
								}
								if (Double.parseDouble(post_saleprice) < 1) {
									post_saleprice = "0.99";
								}
								
								post_stock = "500";
								post_image = sheet.getCell(13, i).getContents();
				    		  
								// Find Category ID from term_taxnonomy_id in wp_term_taxonomy table
								if (sheet.getCell(1, i).getContents().equals("Fragrance") && sheet.getCell(2, i).getContents().equals("Bath & Body")) {
									post_category = "33";	  
								} else if (sheet.getCell(1, i).getContents().equals("Skincare")) {
									post_category = "23";	 
								} else if (sheet.getCell(1, i).getContents().equals("Haircare")) {
									post_category = "22";	 
								} else if (sheet.getCell(1, i).getContents().equals("Aromatherapy")) {
									post_category = "20";	 
								} else if (sheet.getCell(1, i).getContents().equals("Candle")) {
									post_category = "21";	 
								}
	
		    		  
		    		  // Retrieve post ID to see if the product already exists in database
				      String post_id = db.executeQuery("select post_id from wp_postmeta where meta_key='_sku' and meta_value = '"+sheet.getCell(0, i).getContents()+"'");
				      
				      if (post_id == null || post_id == "") {
					      System.out.println("Inserting Data for SKU " + sheet.getCell(0, i).getContents()+", i.e. Line Number " +i);
					      
					      //Insert into Posts table
					      sql_posts += "('1', '"+datetimestamp+"', '"+datetimestamp+"', '"+ post_content +"', '"+ toCamelCase(post_title) +"', '"+ post_excerpt +"', 'publish', 'open', 'open', '', '"+ post_name +"', '', '', '"+datetimestamp+"', '"+datetimestamp+"', '', '0', 'TEST', '0', 'product', '', '0' )";
	     				  db.executeQuery(sql_posts);
					      String addpost_id = db.executeQuery("select id from wp_posts where post_name = '"+post_name+"'");
		
					      //Insert into PostsMeta table				      
					      if (!post_regularprice.equals("-")) {
					    	  sql_postmeta = "INSERT INTO wp_postmeta (post_id, meta_key , meta_value) Values ("+addpost_id+",'_sku', "+post_sku+"),("+addpost_id+",'_stock', '500'),("+addpost_id+",'_tax_status', 'taxable'),("+addpost_id+",'_backorders', 'no'),("+addpost_id+",'_downloadable', 'no'),("+addpost_id+",'_virtual', 'no'),("+addpost_id+",'_visibility', 'visible'),("+addpost_id+",'_stock_status', 'instock'),("+addpost_id+",'_manage_stock', 'yes'),("+addpost_id+",'_sale_price', "+post_saleprice+"),("+addpost_id+",'_price', "+post_saleprice+"),("+addpost_id+",'image_link1', '"+post_image+"'),("+addpost_id+",'post_product_name', '"+post_product_name+"'),("+addpost_id+",'post_product_type', '"+post_product_type+"'),("+addpost_id+",'brand', '"+post_brand+"'),("+addpost_id+",'_regular_price', "+post_regularprice+"),("+addpost_id+",'_seller', 'fragranceNet')";
					    	  db.executeQuery(sql_postmeta);
						  } else {
							  sql_postmeta = "INSERT INTO wp_postmeta (post_id, meta_key , meta_value) Values ("+addpost_id+",'_sku', "+post_sku+"),("+addpost_id+",'_stock', '500'),("+addpost_id+",'_tax_status', 'taxable'),("+addpost_id+",'_backorders', 'no'),("+addpost_id+",'_downloadable', 'no'),("+addpost_id+",'_virtual', 'no'),("+addpost_id+",'_visibility', 'visible'),("+addpost_id+",'_stock_status', 'instock'),("+addpost_id+",'_manage_stock', 'yes'),("+addpost_id+",'_sale_price', "+post_saleprice+"),("+addpost_id+",'_price', "+post_saleprice+"),("+addpost_id+",'image_link1', '"+post_image+"'),("+addpost_id+",'post_product_name', '"+post_product_name+"'),("+addpost_id+",'post_product_type', '"+post_product_type+"'),("+addpost_id+",'brand', '"+post_brand+"'),("+addpost_id+",'_seller', 'fragranceNet')";			
							  db.executeQuery(sql_postmeta);
						  }
					      
					      //Insert into Categories table				      						
					      sql_postmeta = "INSERT INTO wp_term_relationships (object_id, term_taxonomy_id , term_order) Values ("+addpost_id+",'"+post_category+"', '0')"; 
					      db.executeQuery(sql_postmeta);
						  
				      } else {
					      System.out.println("Updating Data for SKU " + sheet.getCell(0, i).getContents()+", i.e. Line Number " +i);		
					      
					      // Remove once we fix all URL Issues
				    	  sql_postmeta = "UPDATE wp_posts SET post_name = '"+post_name+"' WHERE id = '"+post_id+"'"; 
	     				  db.executeQuery(sql_postmeta);	
	     				  
					      //Update PostsMeta table				      
					      if (!post_regularprice.equals("-")) {
							  	sql_postmeta = "UPDATE wp_postmeta SET meta_value= CASE WHEN meta_key='_regular_price' THEN "+post_regularprice+" WHEN meta_key='_sale_price' THEN "+post_saleprice+" WHEN meta_key='_price' THEN  "+post_saleprice+" WHEN meta_key='_stock' THEN "+post_stock+"  END WHERE post_id = "+post_id+" AND meta_key IN ('_regular_price' ,  '_sale_price', '_price', '_stock')"; 
							  	db.executeQuery(sql_postmeta);	
						  } else {
							  	sql_postmeta = "UPDATE wp_postmeta SET meta_value= CASE WHEN meta_key='_sale_price' THEN "+post_saleprice+" WHEN meta_key='_price' THEN  "+post_saleprice+" WHEN meta_key='_stock' THEN "+post_stock+"  END WHERE post_id = "+post_id+" AND meta_key IN ('_regular_price' , '_sale_price', '_price', '_stock')"; 
							  	db.executeQuery(sql_postmeta);	
						  }
					
				      }
    			  }	  
    		  }
		   }
			  
	    } catch (BiffException e) {
	    	e.printStackTrace();
	    }		
	    
		db.closeDB();
	}
	
	static String toCamelCase(String s){
	   s = s.replace("    "," ").replace("   "," ").replace("  ", " ");
	   String[] parts = s.split(" ");
	   String camelCaseString = "";
	   for (String part : parts){
		   if (part.equals("EDT")) {
			   camelCaseString = camelCaseString + " " + part;
		   } else {
			   camelCaseString = camelCaseString + " " + toProperCase(part);			   
		   }
		}
	   return camelCaseString;
	}

	static String toProperCase(String s) {
		return s.substring(0, 1).toUpperCase() +
		s.substring(1).toLowerCase();
	}
}
