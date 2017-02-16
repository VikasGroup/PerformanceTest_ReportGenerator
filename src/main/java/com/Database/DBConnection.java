package com.Database;

import java.beans.Statement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import com.Excel.PropertyReader;

public class DBConnection{
	public static Connection conn() throws ClassNotFoundException, SQLException{
		 // JDBC driver name and database URL
		PropertyReader prop=new PropertyReader();
		System.out.println("jknjkn"+prop.propp.getProperty("JDBC_DRIVER"));
		    String JDBC_DRIVER = prop.propp.getProperty("JDBC_DRIVER"); 
		    String DB_URL = prop.propp.getProperty("DB_URL"); 

		   //  Database credentials
		    String USER = prop.propp.getProperty("USER"); 
		    String PASS = prop.propp.getProperty("PASS"); 
		   
		  
		   Connection conn = null;
		  
		  
		      //STEP 2: Register JDBC driver
		      Class.forName("com.mysql.jdbc.Driver");

		      //STEP 3: Open a connection
		     // System.out.println("Connecting to database...");
		      conn = DriverManager.getConnection(DB_URL,USER,PASS);
		      return conn;
	}
}
