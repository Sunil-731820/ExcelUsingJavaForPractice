package com.java.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class WriteDataToExcel {
	public static void main(String[] args) throws IOException {
		
		 HSSFWorkbook workbook = new HSSFWorkbook();
		  
	        // spreadsheet object
	        HSSFSheet spreadsheet
	            = workbook.createSheet(" Student Data ");
	  
	        // creating a row object
	        HSSFRow row;
	  
	        // This data needs to be written (Object[])
	        Map<String, Object[]> studentData
	            = new TreeMap<String, Object[]>();
	  
	        studentData.put(
	            "1",
	            new Object[] { "Roll No", "NAME", "Year" });
	  
	        studentData.put("2", new Object[] { "128", "Aditya",
	                                            "2nd year" });
	  
	        studentData.put(
	            "3",
	            new Object[] { "129", "Narayana", "2nd year" });
	  
	        studentData.put("4", new Object[] { "130", "Mohan",
	                                            "2nd year" });
	  
	        studentData.put("5", new Object[] { "131", "Radha",
	                                            "2nd year" });
	  
	        studentData.put("6", new Object[] { "132", "Gopal",
	                                            "2nd year" });
	        
	        studentData.put(
		            "7",
		            new Object[] { "Roll No", "NAME", "Year" });
		  
		        studentData.put("8", new Object[] { "128", "Aditya",
		                                            "2nd year" });
		  
		        studentData.put(
		            "9",
		            new Object[] { "129", "Narayana", "2nd year" });
		  
		        studentData.put("10", new Object[] { "130", "Mohan",
		                                            "2nd year" });
		  
		        studentData.put("11", new Object[] { "131", "Radha",
		                                            "2nd year" });
		  
		        studentData.put("12", new Object[] { "132", "Gopal",
		                                            "2nd year" });
		        studentData.put("13", new Object[] { "132", "Gopal",
                "2nd year" });
		        studentData.put("14", new Object[] { "132", "Gopal",
                "2nd year" });
		        studentData.put("15", new Object[] { "132", "Gopal",
                "2nd year" });
		        studentData.put("16", new Object[] { "132", "Gopal",
                "2nd year" });
		        
		        studentData.put("17", new Object[]  {"132","Sunil","4 Year"});
		        
	  
	        Set<String> keyid = studentData.keySet();
	  
	        int rowid = 0;
	  
	        // writing the data into the sheets...
	  
	        for (String key : keyid) {
	  
	            row = spreadsheet.createRow(rowid++);
	            Object[] objectArr = studentData.get(key);
	            int cellid = 0;
	  
	            for (Object obj : objectArr) {
	                Cell cell = row.createCell(cellid++);
	                cell.setCellValue((String)obj);
	            }
	        }
	  
	        // .odf is the format for libre office excel format Excel Sheets...
	        // writing the workbook into the file...
	        FileOutputStream out = new FileOutputStream(
	            new File("javatpoint1.odf"));
	  
	        workbook.write(out);
	        out.close();
		
	}

}
