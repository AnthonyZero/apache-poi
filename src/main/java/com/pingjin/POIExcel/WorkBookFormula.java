package com.pingjin.POIExcel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;

public class WorkBookFormula {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileOutputStream out=new FileOutputStream(new File("F://demo3.xls"));
		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("表1");
		HSSFRow row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		cell.setCellValue(2);
		cell=row.createCell(1);
		cell.setCellValue(3);
		cell=row.createCell(3);
		cell.setCellFormula("sum(A1:B1)");
		//数学公式
		cell=row.createCell(4);
		cell.setCellFormula("max(a1:b1)");
		cell=row.createCell(5);
		cell.setCellFormula("min(a1:b1)");
		/*cell=row.createCell(6);
		cell.setCellFormula("power(a1:b1)");*/
		cell=row.createCell(7);
		cell.setCellFormula("count(a1:b1)");
		cell=row.createCell(8);
		cell.setCellFormula("PRODUCT(a1:b1)");
		
		//超链接URL
		row=sheet.createRow(1);
		cell=row.createCell(1);
		CreationHelper creationHelper=workbook.getCreationHelper();
		Hyperlink link=creationHelper.createHyperlink(Hyperlink.LINK_URL);
		link.setAddress("http://www.baidu.com");
		cell.setCellValue("网站链接");
		cell.setHyperlink(link);
		workbook.write(out);
		out.close();
		System.out.println("success");
	}

}
