package com.gun.tm.tool.excel;

import com.google.common.collect.Lists;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.parser.Tag;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zhaolei
 * @create 2016-08-20 14:20
 */
public class Main6 {
    public static void main(String[] args) throws IOException {
        Document doc = Jsoup.parse(new File("d:\\4.html"), "UTF-8");
        Element body = doc.body();
        Elements elements = body.children();
        List<Element> pElementList = Lists.newArrayList();
        String txName="";
        for (Element element : elements) {
            if(element.tag().getName().equals("p")){
                if(isDaTi(element)){
                    txName = getDaTi(element);
                }
                pElementList.add(element);
                if(isDaAn(element) && !txName.equals("")){
                    Element txElement = createElement(txName);
                    pElementList.add(pElementList.size()-1,txElement);
                }
            }
        }

        List<List<Element>> reList = regroup(pElementList);
        for (List<Element> elementList : reList) {
            for (Element element : elementList) {
                System.out.println(element.text());
            }
            System.out.println("=============");
        }
    }

    public static List<List<Element>> regroup(List<Element> pElementList){
        List<List<Element>> returnList = Lists.newArrayList();
        List<Element> tiMuElementList = Lists.newArrayList();
        for (Element element : pElementList) {
            if(isTiGan(element)){
                if(tiMuElementList.size()>0){
                    returnList.add(tiMuElementList);
                }
                tiMuElementList = Lists.newArrayList();
            }
            tiMuElementList.add(element);
        }
        return returnList;
    }

    public static boolean isTiGan(Element element){
        String text = element.text();
        String regEx="^(\\d+(\\.|．)).+";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static boolean isDaTi(Element element){
        String text = element.text();
        String regEx="^(一|二|三|四|五|六|七|八|九|十)、.+";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static String getDaTi(Element element){
        //^[一|二|三|四|五|六|七|八|九|十]、.+
        String name = "";
        String text = element.text();
        String regEx="^(一|二|三|四|五|六|七|八|九|十)、.+";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        String nameText = "";
        while (m.find()){
            nameText = m.group();
        }
        if(!nameText.trim().equals("")){
            name = nameText.substring(nameText.indexOf("、")+1,nameText.indexOf("（")).trim();
        }
        return name;
    }

    public static Element createElement(String elementName){
        Element element = new Element(Tag.valueOf("p"),"");
        element.attr("class","p3");
        element.html("<span class=\"s3\">【题型】</span><span class=\"s5\">"+elementName+"</span>");
        return element;
    }

    public static boolean isDaAn(Element element){
        String text = element.text();
        String regEx="^(〖|【)?(答案|解:|解：)(〗|】)?.+";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static boolean isXuanxiang(Element element){

    }
}
