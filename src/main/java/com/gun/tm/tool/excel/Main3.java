package com.gun.tm.tool.excel;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.util.XMLHelper;
import org.w3c.dom.Document;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

/**
 * @author zhaolei
 * @create 2016-08-16 10:53
 */
public class Main3 {
    public static void main(String[] args) throws Exception {
//        final String path = "F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\";
        final String path = "F:\\中农网\\word导入excel\\样例\\样例3\\湖北襄阳五中八年级数学6.8\\";
//        final String file = "2016年山西省《学习方法报》2016-2017学年第一学期数学沪科八年级第2期《第12章  一次函数（12．1）自我评估》拆解题目数学试卷试卷编号：1111613021508021.doc";
        final String file = "襄阳五中实验中八年级下学期六月月考数学试题.doc";
//        System.out.println("Converting " + args[0]);
//        System.out.println("Saving output to " + args[1]);
        Document doc = process(new File(path+file));
        DOMSource domSource = new DOMSource(doc);
        OutputStream outputStream = new FileOutputStream(new File("F:\\中农网\\5.html"));
        StreamResult streamResult = new StreamResult(outputStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty("encoding", "GB2312");
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
    public static void writeFile(String content, String path) {
        FileOutputStream fos = null;
        BufferedWriter bw = null;
        try {
            File file = new File(path);
            fos = new FileOutputStream(file);
            bw = new BufferedWriter(new OutputStreamWriter(fos,"GB2312"));
            bw.write(content);
        } catch (FileNotFoundException fnfe) {
            fnfe.printStackTrace();
        } catch (IOException ioe) {
            ioe.printStackTrace();
        } finally {
            try {
                if (bw != null)
                    bw.close();
                if (fos != null)
                    fos.close();
            } catch (IOException ie) {
            }
        }
    }
}
