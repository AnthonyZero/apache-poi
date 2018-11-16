package com.pingjin.poiword;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.extractor.WordExtractor;

/**
 * 通过WordExtractor读文件
 * @author Anthony_One
 *
 */
public class HWPFReadDataTest {
	/**
	 * 输出SummaryInfomation
	 * 
	 * @param info
	 */
	private static void printInfo(SummaryInformation info){
		// 作者
		System.out.println(info.getAuthor());
		// 字符统计
		System.out.println(info.getCharCount());
		// 页数
		System.out.println(info.getPageCount());
		// 标题
		System.out.println(info.getTitle());
		// 主题
		System.out.println(info.getSubject());
	}
	
	/**
	 * 输出DocumentSummaryInfomation
	 * 
	 * @param info
	 */
	private static void printInfo(DocumentSummaryInformation info) {
		// 分类
		System.out.println(info.getCategory());
		// 公司
		System.out.println(info.getCompany());
	}
 
	/**
	 * 关闭输入流
	 * 
	 * @param is
	 */
	private static void closeStream(InputStream is) {
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
 
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		FileInputStream in=new FileInputStream(new File("F://demo.doc"));
		WordExtractor wordExtractor=new WordExtractor(in);
		//输出word文档所有的文本 
		System.out.println(wordExtractor.getText());
		System.out.println(wordExtractor.getTextFromPieces());
		System.out.println("-----------------------");
		//输出页眉的内容  
	    System.out.println("页眉：" + wordExtractor.getHeaderText()); 
	    System.out.println("-----------------------");
	    //输出页脚的内容  
	    System.out.println("页脚：" + wordExtractor.getFooterText());
	    System.out.println("-----------------------");
		// 输出当前word文档的元数据信息，包括作者、文档的修改时间等。
		System.out.println(wordExtractor.getMetadataTextExtractor().getText());
		System.out.println("-----------------------");
		// 获取各个段落的文本
		String paraTexts[] = wordExtractor.getParagraphText();
		for (int i = 0; i < paraTexts.length; i++) {
			System.out.println("Paragraph " + (i + 1) + " : " + paraTexts[i]);
		}
		System.out.println("-----------------------");
		// 输出当前word的一些信息
		printInfo(wordExtractor.getSummaryInformation());
		System.out.println("-----------------------");
		// 输出当前word的一些信息
		printInfo(wordExtractor.getDocSummaryInformation());
		System.out.println("-----------------------");
		closeStream(in);
	}

}
