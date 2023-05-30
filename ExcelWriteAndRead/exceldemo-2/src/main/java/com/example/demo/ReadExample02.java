package com.example.demo;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


//讀取Excel 檔案程式範例
@SpringBootApplication
public class ReadExample02 {

	public static void main(String[] args) {
		FileInputStream input = null;
		String fileName = "file.xlsx";
		
		try {
			input = new FileInputStream(fileName);
			
			@SuppressWarnings("resource")
			XSSFWorkbook book = new XSSFWorkbook(input);
			XSSFSheet sheet = book.getSheetAt(0);
            XSSFRow row = sheet.getRow(0);
             
            XSSFCell cell = null;
            String title = "讀取單元格內容: ";
            for(int i = 0; i< 3; i++) {
                cell = row.getCell(i);
                 
                if (cell.getCellType() == CellType.STRING) {
                    System.out.println(title + cell.getStringCellValue());
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    System.out.println(title + cell.getNumericCellValue());
                }
		}
	}catch (Exception e) {
        e.printStackTrace();
    }
	}

}
