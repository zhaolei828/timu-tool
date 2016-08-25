package com.gun.tm.tool.excel;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.util.XMLHelper;
import org.w3c.dom.Document;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;

/**
 * @author zhaolei
 * @create 2016-08-16 10:53
 */
public class Main3 {
    public static void main(String[] args) throws Exception {
//        final String path = "F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\";
        final String path = "F:\\项目\\其他\\转excel\\比多重复试题的样例\\湘教七年级\\";
//        final String file = "2016年山西省《学习方法报》2016-2017学年第一学期数学沪科八年级第2期《第12章  一次函数（12．1）自我评估》拆解题目数学试卷试卷编号：1111613021508021.doc";
        final String file = "《学习方法报》2016-2017学年第一学期数学湘教七年级第5期《 有理数（1.5）自我评估》拆解题目试卷编号：1111613022607051.doc";
//        System.out.println("Converting " + args[0]);
//        System.out.println("Saving output to " + args[1]);
        Document doc = process(new File(path+file));
        DOMSource domSource = new DOMSource(doc);
        StreamResult streamResult = new StreamResult(new File("d:\\5.html"));
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty("encoding", "UTF-8");
        serializer.setOutputProperty("indent", "yes");
        serializer.setOutputProperty("method", "html");
        serializer.transform(domSource, streamResult);
    }
    static Document process(File docFile) throws Exception {
        HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(docFile);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument());
        wordToHtmlConverter.processDocument(wordDocument);
        return wordToHtmlConverter.getDocument();
    }
}
