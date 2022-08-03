package com.java.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class NewSheetExcel {
	public static void main(String[] args) throws IOException {
		Workbook book = new HSSFWorkbook();
		try {
			OutputStream fileOut = new FileOutputStream("javatpoint.ods");
			Sheet sheet1 = book.createSheet("firstsheet");
			Sheet sheet2 = book.createSheet("secondsheet");
			book.write(fileOut);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
//			e.printStackTrace();
			System.out.println(e.getMessage());
		}
		
	}

}
