package com.pingjin.context;

import com.pingjin.strategy.WordHtmlConverter;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import java.io.File;
import java.io.IOException;

/**
 * 上下文 (策略模式)
 */
public class ConverterContext {

    private WordHtmlConverter converter;

    public ConverterContext(WordHtmlConverter wordHtmlConverter) {
        this.converter = wordHtmlConverter;
    }

    /**
     * 文档文件转化为HTML文件
     * @param source word(源文件)
     * @param storage 存储目录
     * @param target  html(目标文件名称)
     * @throws ParserConfigurationException
     * @throws TransformerException
     * @throws IOException
     */
    public void convertToHtml(File source, String storage, String target) throws ParserConfigurationException, TransformerException, IOException {
        converter.convert(source, storage, target);
    }
}
