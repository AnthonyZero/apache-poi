package com.pingjin.ReadXmlFile;

import java.io.File;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

public class XMLReader {
	
	public static void main(String args[]) throws Exception{
		File file=new File("D:/Workspaces/EclipseProjects/ApachePOI/src/main/resources/config.xml");
		if(!file.exists()){
			System.out.println("文件不存在");
		}else{
			SAXReader reader=new SAXReader();
			Document document=reader.read(file);
			Element root=document.getRootElement();
			Element itr=(Element) root.elements("localPath");
			System.out.println(itr==null);
			String path=itr.elementText("localPath").trim();
			System.out.println(path);
		}
	}
}
