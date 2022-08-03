package com.java.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class CreateWorkBook {
	
	public static void main(String[] args) throws IOException {
		System.out.println("This is my first excel sheet");
		Workbook book = new HSSFWorkbook();
		try {
			OutputStream fileName = new FileOutputStream("javatpoint.xls");
			book.write(fileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
//			e.printStackTrace();
			System.out.println(e.getMessage());
		}
	}

}
