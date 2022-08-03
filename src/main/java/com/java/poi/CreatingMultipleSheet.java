package com.java.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

public class CreatingMultipleSheet {
	public static void main(String[] args) throws IOException {
		System.out.println("creating the multiple sheet using java");
		HSSFWorkbook book = new HSSFWorkbook();
		 // An output stream accepts output bytes and
        // sends them to sink
        OutputStream fileOut = new FileOutputStream("Geeks.ods");
        //Creating the sheet using createsheet()
        
        Sheet sheet1 = book.createSheet("Array");
        Sheet sheet2 = book.createSheet("Dynamic Programming");
        Sheet sheet3 = book.createSheet("String");
        Sheet sheet4 = book.createSheet("Bit Manipulations");
        Sheet sheet5 = book.createSheet("Gready problems");
        
        System.out.println("now check the Numnber of the sheet available ");
        book.write(fileOut);
        
        int NumberOfSheet = book.getNumberOfSheets();
        System.out.println("the Number of the sheet is " + NumberOfSheet);
        
        
        
        
	}

}
