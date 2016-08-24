package com.pingjin.POIExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class UpdateWorkBookData {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		//先获取输入流获得数据
		File file=new File("F://demo.xls");
		FileInputStream in=new FileInputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		//修改单元格数据
		HSSFSheet sheet=workbook.getSheet("table sheet");
		sheet.getRow(0).getCell(1).setCellValue("修改数据");
		//输出流保存
		FileOutputStream out=new FileOutputStream(file);
		workbook.write(out);
		System.out.println("修改成功");
	}

}
