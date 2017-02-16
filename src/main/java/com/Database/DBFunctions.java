package com.Database;




import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;



public class DBFunctions {
	public static double [] pageTimes=new double[3]  ;
	public double getPgLoadTime(String tableName,String date,String pgName,String time) throws ClassNotFoundException, SQLException {
		 Statement stmt = null;
		 Connection conn=DBConnection.conn();
		 int x=0;
		double loadtime=0;
		stmt =  conn.createStatement();
		
	      String sql;
	      sql = "SELECT * FROM `"+tableName+"` WHERE daystamp='"+date+"' and timestamp='"+time+"'";
	      //System.out.println(sql);
	      ResultSet rs = stmt.executeQuery(sql);

	      
	      //STEP 5: Extract data from result set
	      while(rs.next()){
	         //Retrieve by column name
	         double pageTime  = rs.getDouble(pgName);
	         pageTimes[x]=pageTime;
	         loadtime=loadtime+pageTime;
	         //Display values
	       //  System.out.println("pageLoadTime: " + pageTime);
	       x++;
	      }
	     
	      //STEP 6: Clean-up environment
	      rs.close();	      
	      conn.close();
	      loadtime=loadtime/3;
	      double val=Math.round(loadtime*100);
	      return val/100;
	}
		public int getCount(String tableName) throws SQLException, ClassNotFoundException{
			int rowcount=0;
		Statement stmt = null;
			 Connection conn=DBConnection.conn();
			double loadtime=0;
			stmt =  conn.createStatement();
		      String sql;
		      sql = "SELECT COUNT(*) as rowcount FROM `"+tableName+"`;";
		      //System.out.println(sql);
		      ResultSet rs = stmt.executeQuery(sql);

		      
		      //STEP 5: Extract data from result set
		      while(rs.next()){
		         //Retrieve by column name
		         rowcount  = rs.getInt("rowcount");
		         //Display values
		        // System.out.println("rowcount: " + rowcount);
		       
		      }
		     
		      //STEP 6: Clean-up environment
		      rs.close();	      
		      conn.close();
		      return rowcount;
		}
	
}
