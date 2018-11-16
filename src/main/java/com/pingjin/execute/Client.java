package com.pingjin.execute;

import com.pingjin.context.ConverterContext;
import com.pingjin.strategy.WordHtmlConverter;
import com.pingjin.strategy.achieve.Word03ToHtmlAchieve;
import com.pingjin.strategy.achieve.Word07ToHtmlAchieve;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.File;
import java.io.IOException;

public class Client {

    public static void main(String[] args) {
        File source = new File("F:/demo.docx");
        if (!source.exists() || !source.isFile()) {
            System.out.println("file not exist");
            return;
        }
        WordHtmlConverter wordHtmlConverter = null;
        if (source.getPath().endsWith("doc") || source.getPath().endsWith("DOC")) {
            wordHtmlConverter = new Word03ToHtmlAchieve();
        } else if (source.getPath().endsWith("docx") || source.getPath().endsWith("DOCX")) {
            wordHtmlConverter = new Word07ToHtmlAchieve();
        } else {
            System.out.println("word file suffix not support");
            return;
        }
        ConverterContext context = new ConverterContext(wordHtmlConverter);
        try {
            context.convertToHtml(source, "E:/filedata", "index.html");
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
