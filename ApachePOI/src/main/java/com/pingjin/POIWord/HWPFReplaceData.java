package com.pingjin.POIWord;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.junit.Test;

public class HWPFReplaceData {

	/**
	 * 通过模板word替换指定的内容
	 * @throws Exception 
	 */
	@Test
	public void testWrite() throws Exception{
		FileInputStream in=new FileInputStream(new File("F:\\demo.doc"));
		HWPFDocument doc=new HWPFDocument(in);
		Range range=doc.getRange();
		// 把range范围内的${reportDate}替换为当前的日期
		range.replaceText("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
		range.replaceText("apple", "100.00");
		range.replaceText("banana", "200.00");
		range.replaceText("total", "300.00");
		OutputStream os = new FileOutputStream("F:\\demo.doc");
		// 把doc输出到输出流中
		doc.write(os);
		in.close();
		os.close();
	}
}
