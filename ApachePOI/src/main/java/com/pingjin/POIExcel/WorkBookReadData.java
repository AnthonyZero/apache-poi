package com.pingjin.POIExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class WorkBookReadData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file=new File("F://测试.xls");
		FileInputStream in=new FileInputStream(file);
		HSSFWorkbook workbook=new HSSFWorkbook(in);
		HSSFSheet sheet1=workbook.getSheetAt(0);
		for(Row row:sheet1){
			for(Cell cell:row){
				System.out.println(cell.getRowIndex()+" "+cell);
			}
			System.out.println();
		}
	}

}
