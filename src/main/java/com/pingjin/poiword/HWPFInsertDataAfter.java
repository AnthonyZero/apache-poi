package com.pingjin.poiword;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class HWPFInsertDataAfter {
	
	public static void main(String args[]) throws Exception{
		FileInputStream in=new FileInputStream(new File("F:/测试文件.docx"));
		XWPFDocument doc=new XWPFDocument(in);
		List<XWPFParagraph> _paraList=doc.getParagraphs();
		// 创建一个段落
		/*XWPFParagraph para = doc.createParagraph();*/
		// 一个XWPFRun代表具有相同属性的一个区域。
		/*XWPFRun run = para.createRun();
		run.setBold(true); // 加粗
		run.setText("加粗的内容");
		run.setColor("FF0000");
		run.setText("红色的字。");*/
		XWPFParagraph lastPara=_paraList.get(_paraList.size()-1);
		List<XWPFRun> runs=lastPara.getRuns();
		System.out.println(runs.size());
		for(int i=0;i<runs.size();i++){
			lastPara.removeRun(i);
		}
		XWPFRun newRun=lastPara.createRun();
		newRun.setText("他一天允许重复发到付");
		newRun.setColor("FF0000");
		newRun.setBold(true);
		newRun.setFontSize(14);
		/*for(XWPFParagraph parag:_paraList){
			System.out.println("段落"+i+":"+parag.getText());
			i++;
		}*/
		/*System.out.println(_paraList.get(1).getText());
		List<XWPFRun> runs=_paraList.get(1).getRuns();
		System.out.println(runs.size());*/
		/*runs.get(0).setText("我的末尾");*/
		FileOutputStream out=new FileOutputStream("F:/测试文件.docx");
		doc.write(out);
		in.close();
		out.close();
		System.out.println("success");
		/*CharacterProperties cp = new CharacterProperties();
        cp.setBold(true);
        cp.setCharacterSpacing(10);
        cp.setChse(cp.SPRM_CHARSCALE);
        cp.setCapitalized(true);*/
		/*int numPara=range.numParagraphs();
		for (int i = 0; i < numPara; i++) {
			System.out.println("段落" + (i + 1) + "：" + range.getParagraph(i).text());
			
			if(i==(numPara-1)){
				range.getParagraph(i).insertAfter("xxxx");
			}
		}*/
		
		/*int numSec=range.numSections();
		for(int i=0;i<numSec;i++){
			System.out.println("小节"+(i+1)+":"+range.getSection(i).text());
			range.getSection(i).insertAfter("我的评语");
		}*/
		
	}
}
