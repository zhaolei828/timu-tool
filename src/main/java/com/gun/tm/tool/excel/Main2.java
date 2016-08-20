package com.gun.tm.tool.excel;

import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author zhaolei
 * @create 2016-08-15 18:33
 */
public class Main2 {
    public static void main(String[] args) throws IOException {
//        String root = "target";
//        String fileOutName = root + "/" + fileInName + ".html";

//        long startTime = System.currentTimeMillis();

        XWPFDocument document = new XWPFDocument(new FileInputStream(new File("F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\4.docx")));

        XHTMLOptions options = XHTMLOptions.create();// .indent( 4 );
        // Extract image
        File imageFolder = new File("F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\images2\\");
        options.setExtractor(new FileImageExtractor(imageFolder ));
        // URI resolver
        options.URIResolver(new FileURIResolver(imageFolder));

        OutputStream out = new FileOutputStream( new File( "F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例1\\沪科八年级\\4.html" ) );
        XHTMLConverter.getInstance().convert( document, out, options );
    }
}
