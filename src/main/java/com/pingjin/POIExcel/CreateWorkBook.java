package com.pingjin.POIExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkBook {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		HSSFWorkbook workbook=new HSSFWorkbook();
		FileOutputStream out=new FileOutputStream(new File("F://示例.xls"));
		workbook.write(out);
		out.close();
		System.out.println("xls文档创建成功");
		/*File file=new File("F://xx.xls");
		FileInputStream in=new FileInputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		if(file.isFile()&&file.exists()){
			System.out.println("文件打开成功");
		}else{
			System.out.println("文件打卡失败");
		}*/
	}

}
