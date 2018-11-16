package com.pingjin.poiexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class WorkBookCellStyle {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file=new File("F://demo.xls");
		//输出流
		FileOutputStream out=new FileOutputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("table sheet");
		HSSFRow row1=sheet.createRow(0);
		HSSFRow row2=sheet.createRow(1);
		
		HSSFCellStyle cellStyle1=workbook.createCellStyle();
		cellStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFCellStyle cellStyle2=workbook.createCellStyle();
		cellStyle2.setBottomBorderColor(HSSFColor.RED.index);
		cellStyle2.setBorderRight(HSSFCellStyle.BORDER_DOTTED);
		HSSFCellStyle style3=workbook.createCellStyle();
	    style3.setFillForegroundColor(HSSFColor.BLUE.index);
	    style3.setFillPattern(HSSFCellStyle.LEAST_DOTS);
	    style3.setAlignment(HSSFCellStyle.ALIGN_FILL);
		
		row1.createCell(0).setCellValue("居中");
		row1.createCell(1).setCellValue("我很好");
		row1.createCell(2).setCellValue("哈哈");
		row1.setHeight((short)800);
		row1.getCell(0).setCellStyle(cellStyle1);
		
		row2.createCell(0).setCellValue("很好");
		row2.getCell(0).setCellStyle(cellStyle2);
		row2.createCell(1).setCellValue("阿斯顿");
		row2.getCell(1).setCellStyle(style3);
		row2.setHeight((short)666);
		workbook.write(out);
		System.out.println("成功");
	}

}
