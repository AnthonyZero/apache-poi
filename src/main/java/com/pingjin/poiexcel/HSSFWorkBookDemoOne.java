package com.pingjin.poiexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HSSFWorkBookDemoOne {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file=new File("F://测试.xls");
		FileOutputStream out=new FileOutputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook();
		
		HSSFSheet sheet=workbook.createSheet("sheet name");
		HSSFRow row=sheet.createRow(1);
		row.createCell(0).setCellValue("平进");
		row.createCell(1).setCellValue(152.632);
		row.createCell(2).setCellValue(false);
		row.createCell(4).setCellValue(new Date());
		
		HSSFRow row2=sheet.createRow(2);
		row2.createCell(0).setCellValue("张三");
		row2.createCell(1).setCellValue(152.62);
		row2.createCell(2).setCellValue(true);
		row2.createCell(4).setCellValue("特色");
		
		workbook.createSheet("表格二");
		workbook.write(out);
		System.out.println("文档创建成功");
	}

}
