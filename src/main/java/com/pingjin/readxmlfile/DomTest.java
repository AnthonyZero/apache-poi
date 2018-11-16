package com.pingjin.readxmlfile;

import java.io.InputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class DomTest {

	public static void main(String[] args) throws Exception{
		InputStream in=ClassLoader.getSystemResourceAsStream("config.xml");
		DocumentBuilderFactory factory=DocumentBuilderFactory.newInstance();
		DocumentBuilder documentBuilder=factory.newDocumentBuilder();
		/*Document document=documentBuilder.parse("D:/Workspaces/EclipseProjects/ApachePOI/src/main/resources/config.xml");*/
		Document document=documentBuilder.parse(in);
		NodeList list=document.getElementsByTagName("localPath");
		Node node=list.item(0);
		System.out.println(node.getTextContent());
	}
}
