package com.pingjin.strategy.achieve;

import com.pingjin.strategy.WordHtmlConverter;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.UUID;

/**
 * 2007版本的word转换为html(.doc后缀结尾)
 */
public class Word07ToHtmlAchieve implements WordHtmlConverter{

    /**
     * word转换为html
     * @param source word(源文件)
     * @param storage 存储目录
     * @param target html(目标文件名称)
     */
    public void convert(File source, String storage, String target) throws IOException {
        File storageCatalog = new File(storage);
        if (!storageCatalog.exists()) {
            storageCatalog.mkdirs();
        }
        // 1加载word文档
        InputStream inputStream = new FileInputStream(source);
        XWPFDocument document = new XWPFDocument(inputStream);

        // 2图片存储 每个文档中的图片存储对应的图片文件夹需不一致 不然后面的会覆盖前面的图片
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        String dateString = formatter.format(new Date());
        File imageFolderFile = new File(storage, dateString + "/" + UUID.randomUUID().toString().replaceAll("-",""));
        XHTMLOptions options = XHTMLOptions.create().URIResolver(
                new FileURIResolver(imageFolderFile));
        options.setExtractor(new FileImageExtractor(imageFolderFile));

        // 3文档转换
        OutputStream out = new FileOutputStream(new File(storage, target));
        XHTMLConverter.getInstance().convert(document, out, options);
    }
}
