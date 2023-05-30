package com.example.demo;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//建立寫入Excel 程式範例，含字型、顏色、框線、自動換行、內容置中及合併儲存格
public class WriterExample03 {

	/**
     * 時間序號
     */
    private static String getTimeNumber() {
        String pattern = "yyyyMMddHHmmssSSS";
        SimpleDateFormat d = new SimpleDateFormat(pattern);
        return d.format(new Date());
    }
 
    public static void main(String[] args) {
 
        @SuppressWarnings("resource")
        XSSFWorkbook book = new XSSFWorkbook();
 
        Font titlefont = book.createFont();
        titlefont.setColor(IndexedColors.BLACK.getIndex());// 顏色
        titlefont.setBold(true); // 粗體
 
        CellStyle style01 = book.createCellStyle();
        style01.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());// 填滿顏色
        style01.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style01.setFont(titlefont);// 設定字體
        style01.setAlignment(HorizontalAlignment.CENTER);// 水平置中
        style01.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直置中
 
        // 設定框線
        style01.setBorderBottom(BorderStyle.THIN);
        style01.setBorderTop(BorderStyle.THIN);
        style01.setBorderLeft(BorderStyle.THIN);
        style01.setBorderRight(BorderStyle.THIN);
        style01.setWrapText(true);// 自動換行
 
        CellStyle style02 = book.createCellStyle();
        style02.setAlignment(HorizontalAlignment.CENTER);// 水平置中
        style02.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直置中
        style02.setBorderBottom(BorderStyle.THICK);
        style02.setBorderTop(BorderStyle.THICK);
        style02.setBorderLeft(BorderStyle.THICK);
        style02.setBorderRight(BorderStyle.THICK);
        style02.setWrapText(true);// 自動換行
 
        XSSFSheet sheet1 = book.createSheet("工作表1");
 
        XSSFRow titlerow = sheet1.createRow(0);
        for (int i = 0; i < 6; i++) {
            XSSFCell cell = titlerow.createCell(i);
            cell.setCellStyle(style01);
            cell.setCellValue("標題 Cell 0 " + i);
            // sheet1.autoSizeColumn(i); // 自動調整欄位寬度
        }
 
        for (int x = 1; x < 10; x++) {
            XSSFRow row = sheet1.createRow(x);
 
            for (int y = 0; y < 6; y++) {
                XSSFCell cell = row.createCell(y);
                cell.setCellStyle(style02);
 
                cell.setCellValue("中文 Cell " + x + " " + y);
                // sheet1.autoSizeColumn(y); // 自動調整欄位寬度
            }
 
            XSSFCell cell = row.createCell(5);
            cell.setCellValue(100);
            cell.setCellStyle(style02);
            // sheet1.autoSizeColumn(5);
        }
 
        XSSFSheet sheet2 = book.createSheet("工作表2");
        for (int x = 0; x < 5; x++) {
            XSSFRow row = sheet2.createRow(x);
 
            for (int y = 0; y < 5; y++) {
                XSSFCell cell = row.createCell(y);
 
                if (x == 0) {
                    cell.setCellStyle(style01);
                } else {
                    cell.setCellStyle(style02);
                }
 
                cell.setCellValue("中文 title " + x + " " + y);
 
                // sheet2.autoSizeColumn(y); // 自動調整欄位寬度
            }
        }
 
        // 指定檔案名稱
        String fileName = "Test_%1$s.xlsx";
        fileName = String.format(fileName, getTimeNumber());
 
        /*
         * 尚未指定檔案路徑，檔案建立在本執行專案內 儲存工作簿
         */
        try (FileOutputStream os = new FileOutputStream(fileName)) {
            book.write(os);
            System.out.println(fileName + " excel export finish.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
 
}
