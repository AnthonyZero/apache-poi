package com.pingjin.strategy;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * 转换器
 */
public interface WordHtmlConverter {

    /**
     * Word转HTML
     * @param source word(源文件)
     * @param storage 存储目录
     * @param target html(目标文件名称)
     */
    void convert(File source, String storage, String target) throws IOException, ParserConfigurationException, TransformerException;
}
