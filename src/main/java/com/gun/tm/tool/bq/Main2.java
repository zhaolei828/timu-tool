package com.gun.tm.tool.bq;

import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.RFonts;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.select.Elements;

import java.io.File;
import java.net.URL;

/**
 * @author zhaolei
 * @create 2016-08-13 11:37
 */
public class Main2 {
    public static void main(String[] args) throws Exception {
        Document doc = Jsoup.parse(new File("C:/Users/Administrator/Desktop/muban.html"),"UTF-8");
        Elements links = doc.getElementsByTag("img"); //将link中的地址替换为绝对地址
//        for (Element element : links) {
//            String src = System.getProperty("user.dir")+"/sample-docs/"+element.attr("src");
//            element.attr("src", src);
//        }
        doc.outputSettings()
                .syntax(Document.OutputSettings.Syntax.xml)
                .escapeMode(Entities.EscapeMode.xhtml);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage(PageSizePaper.valueOf("A4"), true); //A4纸，//横版:true

        //configSimSunFont(wordMLPackage); //配置中文字体

        XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordMLPackage);

        wordMLPackage.getMainDocumentPart().getContent().addAll( //导入 xhtml
                xhtmlImporter.convert(doc.html(), doc.baseUri()));

//        wordMLPackage.getMainDocumentPart().addAltChunk(AltChunkType.Html, html.getBytes());

        wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/sample-docs/test00000.docx"));
    }
    static void configSimSunFont(WordprocessingMLPackage wordMLPackage) throws Exception {
        Mapper fontMapper = new IdentityPlusMapper();
        wordMLPackage.setFontMapper(fontMapper);

        String fontFamily = "SimSun";

        URL simsunUrl = Main2.class.getClass().getResource("/org/noahx/html2docx/simsun.ttc"); //加载字体文件（解决linux环境下无中文字体问题）
        PhysicalFonts.addPhysicalFonts(fontFamily, simsunUrl);
        PhysicalFont simsunFont = PhysicalFonts.get(fontFamily);
        fontMapper.put(fontFamily, simsunFont);

        RFonts rfonts = Context.getWmlObjectFactory().createRFonts(); //设置文件默认字体
        rfonts.setAsciiTheme(null);
        rfonts.setAscii(fontFamily);
        wordMLPackage.getMainDocumentPart().getPropertyResolver()
                .getDocumentDefaultRPr().setRFonts(rfonts);
    }
}
