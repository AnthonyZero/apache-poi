package com.pingjin.htmlreplace;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;

public class ReplaceText {

	private static final String older="img src=\"f:";
	private static final String newStr="img src=\"/res";
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 BufferedReader br;
	        try {
	            br = new BufferedReader(new InputStreamReader(new FileInputStream("F://test//test3.html")));
	            StringBuffer sb = new StringBuffer();
	            String str = null;
	            while((str = br.readLine())!= null){
	            	sb.append(str);
	            }
	            FileOutputStream file = new FileOutputStream("F://test//test3.html");
	            file.write(sb.toString().replaceAll(older,newStr).getBytes());
	            br.close();
	            file.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	} 
}

