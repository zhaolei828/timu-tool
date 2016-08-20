package com.gun.tm.tool.excel;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.hwpf.usermodel.Range;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author zhaolei
 * @create 2016-08-15 15:57
 */
public class Main {
    public static void main(String[] args) throws IOException, TransformerException, ParserConfigurationException {
//        FileInputStream in=new FileInputStream(new File("F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\2016年山西省《学习方法报》2016-2017学年第一学期数学沪科八年级第2期《第12章  一次函数（12．1）自我评估》拆解题目数学试卷试卷编号：1111613021508021.doc"));
//        HWPFDocument doc=new HWPFDocument(in);
//        Range rang = doc.getRange();
//        for (int i = 0; i <rang.numParagraphs() ; i++) {
//            Paragraph paragraph = rang.getParagraph(i);
//            System.out.println("paragraph = [" + paragraph + "]");
//        }
        final String path = "F:\\word\\";
        final String file = "04016230208012.doc";
        InputStream input = new FileInputStream(path + file);
        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            public String savePicture(byte[] content, PictureType pictureType,
                                      String suggestedName, float widthInches, float heightInches) {
                return "image\\04016230208012\\"+suggestedName;
//                return suggestedName;
            }
        });
        wordToHtmlConverter.processDocument(wordDocument);
        PicturesTable pTable = wordDocument.getPicturesTable();
        List pics = pTable.getAllPictures();
        if (pics != null) {
            for (int i = 0; i < pics.size(); i++) {
                Picture pic = (Picture) pics.get(i);
                try {
                    String filename = pic.suggestFullFileName();
//                    System.out.println(pic.suggestFullFileName()+": [" + pic.suggestPictureType() + "]");
//                    System.out.println("\tgetMimeType()  = [" + pic.getMimeType() + "]");
                    if (pic.suggestPictureType() == PictureType.UNKNOWN) {
                        filename = filename + ".png";
                    }
                    pic.writeImageContent(new FileOutputStream(path + "image\\04016230208012\\" + filename));
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
//            int length = pics.size();
//            for (int i = 0; i < length; i++) {
//                Range range = new Range(i, i + 1, wordDocument);
//
//                CharacterRun cr = range.getCharacterRun(0);
//                if (pTable.hasPicture(cr)) {
//                    Picture pic = pTable.extractPicture(cr, false);
//                    String afileName = pic.suggestFullFileName();
//                    if (pic.suggestPictureType() == PictureType.UNKNOWN) {
//                        afileName = afileName + ".png";
//                    }
//                    OutputStream out = new FileOutputStream(new File(path + "image\\" + afileName));
//                    pic.writeImageContent(out);
//                }
//            }
        }
        Document htmlDocument = wordToHtmlConverter.getDocument();
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        outStream.close();
        String content = new String(outStream.toByteArray());
        FileUtils.write(new File(path, "04016230208012.html"), content, "utf-8");
    }
}
