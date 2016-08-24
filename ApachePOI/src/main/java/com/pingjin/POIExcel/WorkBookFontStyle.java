package com.pingjin.POIExcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class WorkBookFontStyle {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file=new File("F://demo2.xls");
		FileOutputStream out=null;
		try {
			out=new FileOutputStream(file);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("文件找不到");
			e.printStackTrace();
		}
		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("font style");
		HSSFRow row=sheet.createRow(2);
		HSSFFont font=workbook.createFont();
		font.setFontHeightInPoints((short) 30);
	    font.setFontName("IMPACT");	   
	    font.setItalic(true);	    
	    font.setColor(HSSFColor.BRIGHT_GREEN.index);
	    
	    //把font样式加入到单元格样式中去
	    HSSFCellStyle cellStyle=workbook.createCellStyle();
	    cellStyle.setFont(font);
	    
	    row.createCell(1).setCellValue("你好");
	    row.getCell(1).setCellStyle(cellStyle);
	    
	    workbook.write(out);
	    System.out.println("...");
	}

}
