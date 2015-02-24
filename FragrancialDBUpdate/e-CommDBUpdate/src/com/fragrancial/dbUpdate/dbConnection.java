package com.fragrancial.dbUpdate;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;

import com.mysql.jdbc.ResultSetMetaData;
import com.mysql.jdbc.Statement;

public class dbConnection {

  Connection con = null;

  public void connectDB() throws Exception
  {

    try {
      Class.forName("com.mysql.jdbc.Driver").newInstance();
      con = DriverManager.getConnection("jdbc:mysql://107.191.100.111/fragranc_nallakada", "fragranc_admin", "r#3rd#world#man");
      //con = DriverManager.getConnection("jdbc:mysql://localhost/dev_commerce", "root", "");

      if(!con.isClosed())
        System.out.println("Successfully connected to " + "MySQL server using TCP/IP...");

    } catch(Exception e) {
      System.err.println("Exception: " + e.getMessage());
    } 
  }
  
  
  public String executeQuery(String query) throws Exception {
	  
      Statement stmt = (Statement) con.createStatement();
      String output ="";
      
	  if( stmt.execute(query) == false )
	     {
	      int num = stmt.getUpdateCount() ;
	      //System.out.println( num + " rows affected" ) ;
	     }
	  else
	     {
	      ResultSet         rs = stmt.getResultSet() ;
	      ResultSetMetaData md = (ResultSetMetaData) rs.getMetaData() ;

	      while( rs.next() )
	         {
	          for( int i = 1; i <= md.getColumnCount(); i++ ){
		             //System.out.print( rs.getString(i) + " " ) ;
		          	 output = rs.getString(i);
	          }
	          	 //System.out.println() ;
	         }
	      rs.close() ;
	     }

	  stmt.close() ;
	return output;
  }
	  
  public void closeDB() throws Exception
  {
      try {
        if(con != null)
          con.close();
      } catch(SQLException e) 
      {
    	  
      }
  }	  
	  
}
  
 

