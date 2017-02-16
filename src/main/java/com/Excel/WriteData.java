package com.Excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.Database.DBFunctions;
import com.itextpdf.text.DocumentException;

public class WriteData {
 public void writeToExcel(String [][][] data,String databases[]) throws Exception{
	PropertyReader prop=new PropertyReader();
	int count=0;
	int rowcount=1;
	DBFunctions dbf=new DBFunctions();
	//Create blank workbook
     XSSFWorkbook workbook = new XSSFWorkbook(); 
     //Create a blank sheet
     XSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");
     //Create row object
     XSSFRow row;
     //This data needs to be written (Object[])
     Map < String, Object[] > empinfo = 
     new TreeMap < String, Object[] >();
     empinfo.put( Integer.toString(++count), new Object[] {"Time Stamp", "Plan","Home Page","","","", "Login Page","","","", "Dashbord","","","", "Product Page","","",""});
     empinfo.put( Integer.toString(++count), new Object[] {"", "","1st","2nd","3rd","Average", "1st","2nd","3rd","Average", "1st","2nd","3rd","Average", "1st","2nd","3rd","Average"});
     
     //8,9
 
     for(int y=0;y<rowcount;y++) {
    	 
    	rowcount=dbf.getCount(databases[y])/3;
    	 for(int x=0;x<10;x++){
    				
    	    empinfo.put(Integer.toString(++count) , new Object[] {data[x][y][0], data[x][y][1],data[x][y][15], data[x][y][16], data[x][y][17],data[x][y][5] ,data[x][y][6],data[x][y][7],data[x][y][8],data[x][y][2],data[x][y][9],data[x][y][10],data[x][y][11] ,data[x][y][3],data[x][y][12],data[x][y][13],data[x][y][14],data[x][y][4]});
    	    
    	    
    	 }
    	
    	 empinfo.put(Integer.toString(++count) , new Object[] {"", "","", "", "","" });
    	
     }
     XSSFCellStyle my_style = workbook.createCellStyle();
     int rowid = 0;
     for(int i=1;i<empinfo.size();i++){
    	 Object [] objectArr = empinfo.get(Integer.toString(i));
    	 
    	 row = spreadsheet.createRow(rowid++);
    	 int cellid = 0;
    	 for (Object obj : objectArr)
         {
    		 Cell cell = row.createCell(cellid++);
    		 try{
    			 System.out.println(Double.parseDouble(obj.toString()));
    		 if(Double.parseDouble(obj.toString())>=10){
    			 my_style.setFillPattern(XSSFCellStyle.DIAMONDS);
    			
    			 my_style.setFillForegroundColor(IndexedColors.RED.getIndex());
    			 my_style.setFillBackgroundColor(IndexedColors.RED.getIndex());
    			 cell.setCellStyle(my_style);
    		 }
    		 }
    		 catch(NumberFormatException ex){
    			 
    		 }
             cell.setCellValue((String)obj);
             
    		 	System.out.println(obj);
         }
     }
     for (int i=0; i<10; i++){
    	 spreadsheet.autoSizeColumn(i);
    	}
     
     FileOutputStream out = new FileOutputStream( 
     new File("timereport.xlsx"));
     
     workbook.write(out);
     out.close();
    
     System.out.println( 
     "Writesheet.xlsx written successfully" );
    
 }

}
