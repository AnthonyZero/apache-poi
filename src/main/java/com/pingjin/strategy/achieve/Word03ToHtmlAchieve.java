package com.pingjin.strategy.achieve;

import com.pingjin.strategy.WordHtmlConverter;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 2003版本的word转换为html(.doc后缀结尾)
 */
public class Word03ToHtmlAchieve implements WordHtmlConverter {

    /**
     * word转换为html
     * @param source word(源文件)
     * @param storage 存储目录
     * @param target html(目标文件名称)
     */
    public void convert(File source, final String storage, String target) throws IOException, ParserConfigurationException, TransformerException {
        File storageCatalog = new File(storage);
        if (!storageCatalog.exists()) {
            storageCatalog.mkdirs();
        }
        InputStream inputStream  = new FileInputStream(source);
        HWPFDocument wordDocument = new HWPFDocument(inputStream);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        // 图片存储管理
        SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMdd");
        String dateString = formatter.format(new Date());
        //每个word中的图片对应一个存储图片文件夹
        final File pictureDir = new File(storage, dateString);
        if (!pictureDir.exists()) {
            pictureDir.mkdirs();
        }
        final List<String> pictureFilePaths = new ArrayList<String>();
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
                                      float heightInches) {
                String filePath = pictureDir.getPath() + "/"  + UUID.randomUUID().toString().replaceAll("-","") + "." + pictureType.getExtension();
                pictureFilePaths.add(filePath);
                // 返回每张图片的路径
                return filePath;
            }
        });

        wordToHtmlConverter.processDocument(wordDocument);
        // 存放图片
        List pics = wordDocument.getPicturesTable().getAllPictures();
        if (pics != null && pics.size() != 0) {
            for (int i = 0; i < pics.size(); i++) {
                Picture pic = (Picture) pics.get(i);
                try {
                    if (i < pictureFilePaths.size()) {
                        pic.writeImageContent(new FileOutputStream(pictureFilePaths.get(i)));
                    }
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }

        Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);

        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        StreamResult streamResult = new StreamResult(outStream);

        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();
        String content = new String(outStream.toByteArray());
        FileUtils.write(new File(storage, target), content, "utf-8");
    }
}
