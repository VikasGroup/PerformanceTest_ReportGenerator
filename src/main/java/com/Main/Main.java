package com.Main;

import java.io.IOException;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import com.Database.DBFunctions;

import com.Excel.PropertyReader;
import com.Excel.WriteData;
import com.itextpdf.text.DocumentException;


public class Main {

	public static void main(String[] args) throws Exception {
				
		DBFunctions dbf=new DBFunctions();
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
		Date date = new Date();
		String datenow=dateFormat.format(date)+" "+"00:00:00";
	String [] databases=new String[]{"World_Plan_Pg_Load_Time","Malaysia_Plan_Pg_Load_Time","Philippines_Plan_Pg_Load_Time","Sri_Lanka_Plan_Pg_Load_Time","Turkey_Plan_Pg_Load_Time","India_Plan_Pg_Load_Time","Rwanda_Plan_Pg_Load_Time","Singapore_Plan_Pg_Load_Time","QN_Europe_Plan_Pg_Load_Time","My_Network_Plan_Pg_Load_Time"};
	int rowcount=0;
		
    String [][][] data=new String[10][9][100];
    //9,8
    for(int x=0;x<10;x++){
    	rowcount=dbf.getCount(databases[x])/3;
    	System.out.println(rowcount);
    	for(int y=0;y<rowcount;y++){
    		 data[x][y][0]=Integer.toString(y+7);
    		 data[x][y][1]=databases[x];
 			 data[x][y][2]=Double.toString(dbf.getPgLoadTime(databases[x], datenow ,"loginpage",Integer.toString(y+7)));
 			data[x][y][6]=Double.toString(DBFunctions.pageTimes[0]);
 		   	data[x][y][7]=Double.toString(DBFunctions.pageTimes[1]);
 		    data[x][y][8]=Double.toString(DBFunctions.pageTimes[2]);
 			 
   			 data[x][y][3]=Double.toString(dbf.getPgLoadTime(databases[x], datenow ,"dashboardpage",Integer.toString(y+7)));
   			data[x][y][9]=Double.toString(DBFunctions.pageTimes[0]);
 		   	data[x][y][10]=Double.toString(DBFunctions.pageTimes[1]);
 		    data[x][y][11]=Double.toString(DBFunctions.pageTimes[2]);
   			 
   			 data[x][y][4]=Double.toString(dbf.getPgLoadTime(databases[x], datenow ,"productpage",Integer.toString(y+7)));
   			data[x][y][12]=Double.toString(DBFunctions.pageTimes[0]);
 		   	data[x][y][13]=Double.toString(DBFunctions.pageTimes[1]);
 		    data[x][y][14]=Double.toString(DBFunctions.pageTimes[2]);
   			 
   			 data[x][y][5]=Double.toString(dbf.getPgLoadTime(databases[x], datenow ,"homepage",Integer.toString(y+7)));
   			data[x][y][15]=Double.toString(DBFunctions.pageTimes[0]);
 		   	data[x][y][16]=Double.toString(DBFunctions.pageTimes[1]);
 		    data[x][y][17]=Double.toString(DBFunctions.pageTimes[2]);
   			 
   			
   			
    	}
    	//System.out.println("xxxxxxxxxxxxxxxxx");
    }
   
	WriteData wd=new WriteData();
	wd.writeToExcel(data,databases);
	
	}

}
