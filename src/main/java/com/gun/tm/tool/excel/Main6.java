package com.gun.tm.tool.excel;

import com.google.common.collect.Lists;
import com.gun.tm.tool.model.Timu;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.parser.Tag;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zhaolei
 * @create 2016-08-20 14:20
 */
public class Main6 {
    public static void main(String[] args) throws IOException {
        Document doc = Jsoup.parse(new File("d:\\3.html"), "UTF-8");
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
        List<Timu> timuList = Lists.newArrayList();
        for (List<Element> elementList : reList) {
            if(!isTiGan(elementList.get(0))){
                continue;
            }
            Timu timu = new Timu();
            List<String> xxList = Lists.newArrayList();
            String tigan = "";
            boolean notDaAnFlag = true;
            boolean appenTiGanFlag = false;
            boolean appenDaAnFlag = false;

            StringBuffer tiGanBuffer = new StringBuffer();

            StringBuffer xiaoTiBuffer = new StringBuffer();

            StringBuffer daAnBuffer = new StringBuffer();

            StringBuffer jieXiBuffer = new StringBuffer();
            for (Element element : elementList) {
                System.out.println("[*] "+element.text());
                //题干
                if (isTiGan(element)){
                    tigan = element.text()+"\n";
                    appenTiGanFlag = true;
                    continue;
                }

                //收集题干与小题之间的内容
                if (appenTiGanFlag && !isXiaoTi(element)) {
                    if (element.text().trim().length() > 0) {
                        tiGanBuffer.append(element.text()+"\n");
                    }
                }

                if(isXiaoTi(element) && notDaAnFlag){
                    if (isFirstXiaoTi(element)){
                        xiaoTiBuffer.append(tiGanBuffer);
                    }
                    xiaoTiBuffer.append(element.text()+"\n");
                    appenTiGanFlag = false;
                }

                //选项
                if(isXuanxiang(element)){
                    String[] xx = splitXuanxiang(element);
                    for (String s : xx) {
                        if(!s.trim().equals("")){
                            xxList.add(s.trim());
                        }
                    }
                }
                //答案 or 答案里含解析
                if (isDaAn(element)){
                    String daAn = element.text();
                    daAnBuffer.append(daAn+"\n");
                    notDaAnFlag = false;

                    String[] daAnJieXiArr = daAn.split("(解析|过程|分析)(:|：)?");
                    if (daAnJieXiArr.length>1) {
                        jieXiBuffer.append(daAnJieXiArr[1]);
                    }
                }

                if(isXiaoTi(element) && !notDaAnFlag){
                    daAnBuffer.append(element.text()+"\n");
                }

                //解析
                if (isJieXi(element)){
                    jieXiBuffer.append(element.text().split("(〖|【)?(解析|过程|分析)(:|：)?(〗|】)?")[1]);
                }

                //题型
                if (isTixing(element)) {
                    String tixingText = element.text();
                    String tixing = "";
                    try {
                        tixing = tixingText.substring(tixingText.indexOf("】")+1);
                    }catch (Exception e){

                    }
                    if(tixing.trim().equals("填空题")){
                        timu.setTixing("综合填空");
                    }else if (tixing.trim().equals("选择题")) {
                        timu.setTixing("单选题");
                    }else {
                        timu.setTixing(tixing);
                    }
                }
                //知识点1～5/考点
                if (isZsd(element)) {
                    String zsdText = element.text();
                    String zsd = "";
                    String[] zsdArr = null;
                    try {
                        int endIndex = zsdText.length();
                        if (zsdText.contains("答案")){
                            endIndex = zsdText.indexOf("答案");
                            String daanJieXi = zsdText.split("答案(:|：)?")[1];
                            String[] daanJiexiArr = daanJieXi.split("(解析|过程|分析)(:|：)?");
                            String daan = daanJiexiArr[0];
                            if(daanJiexiArr.length>1){
                                String jiexi = daanJiexiArr[1];
                                jieXiBuffer.append(jiexi);
                            }
                            daAnBuffer.append(daan);
                        }
                        zsd = zsdText.substring(zsdText.indexOf("】")+1,endIndex);
                    }catch (Exception e){

                    }
                    if(zsd.trim().length()>0){
                        if(zsd.contains(" ")){
                            zsdArr = zsd.split(" ");
                        }else if(zsd.contains("；")){
                            zsdArr = zsd.split("；");
                        }else {
                            zsdArr = new String[]{zsd};
                        }
                    }
                    timu.setZsdArr(zsdArr);
                }

                //能力结构
                if (isNengLiJieGou(element)) {
                    String nengliText = element.text();
                    String nengli = "";
                    try {
                        nengli = nengliText.substring(nengliText.indexOf("】")+1);
                    }catch (Exception e){

                    }
                    timu.setNljg(nengli);
                }

                //评价
                if (isPingJia(element)) {
                    String pingJiaText = element.text();
                    String pingjia = "";
                    try {
                        pingjia = pingJiaText.substring(pingJiaText.indexOf("】")+1);
                    }catch (Exception e){

                    }
                    timu.setPingjia(pingjia);
                }
            }
            timu.setTigan(tigan+xiaoTiBuffer.toString());
            timu.setXuanxiang(makeXuanxiang(xxList));
            String[] daAnArr = daAnBuffer.toString().split("(〖|【)?(答案|解)(:|：)?(〗|】)?");
            if (daAnArr.length>1){
                timu.setDaan(daAnArr[1]);
            }else {
                timu.setDaan(daAnArr[0]);
            }

            timu.setJiexi(jieXiBuffer.toString());
            timuList.add(timu);
            System.out.println(timu.getXuanxiang());
            System.out.println("=============");
        }
        writeIntoExcel(timuList);
    }

    public static List<List<Element>> regroup(List<Element> pElementList){
        List<List<Element>> returnList = Lists.newArrayList();
        List<Element> tiMuElementList = Lists.newArrayList();
        for (Element element : pElementList) {
            if(isTiGan(element) || isDaTi(element)){
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
        String regEx="^(\\d+(\\.|．)).+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isXiaoTi(Element element){
        String regEx="^(（|\\()\\d+(）|\\)).+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isFirstXiaoTi(Element element){
        String regEx="^(（|\\()1(）|\\))";
        return isBiaoQian(element,regEx);
    }

    public static boolean isDaTi(Element element){
        String regEx="^(一|二|三|四|五|六|七|八|九|十)、.+";
        return isBiaoQian(element,regEx);
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
        String regEx="^(〖|【)?(答案|解:|解：)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isJieXi(Element element){
        String regEx="^(〖|【)?(解析|过程|分析)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isXuanxiang(Element element){
        String regEx="^(A|B|C|D|E|F)(\\.|．)+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isTixing(Element element){
        String regEx="^(〖|【)?题型(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }
    public static boolean isZsd(Element element){
        String regEx="(〖|【)?\\s*(考点|三级知识点)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isNengLiJieGou(Element element){
        String regEx="^(〖|【)?能力结构(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isPingJia(Element element){
        String regEx="^(〖|【)?难度等级(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static String[] splitXuanxiang(Element element){
        String text = element.text();
        String regEx="(A|B|C|D|E|F)(\\.|．)+";
        String[]xx = text.split(regEx);
        return xx;
    }

    public static String makeXuanxiang(List<String> list){
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < list.size(); i++) {
            switch (i) {
                case 0:
                    sb.append("A::");
                    break;
                case 1:
                    sb.append("B::");
                    break;
                case 2:
                    sb.append("C::");
                    break;
                case 3:
                    sb.append("D::");
                    break;
                case 4:
                    sb.append("E::");
                    break;
                case 5:
                    sb.append("F::");
                    break;
                default:break;
            }
            sb.append(list.get(i)+"\n");
        }
        return sb.toString();
    }

    public static boolean isBiaoQian(Element element,String regex){
        String text = element.text();
        Pattern p = Pattern.compile(regex);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static String getBiaoQianText(Element element){
        String text = element.text();
        String content = "";
        try {
            content = text.substring(text.indexOf("】")+1);
        }catch (Exception e){

        }
        return content;
    }

    public static void  writeIntoExcel(List<Timu> list) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet();


        Font font = wb.createFont();
        font.setFontHeightInPoints((short)14);
        font.setFontName("Courier New");
        font.setBold(true);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);


        String[] head = new String[]{"编号*","学科*","省份","城市","年份","题型*","错误率","题干","备选答案",
                "正确答案","解析"	,"试题评价","典型题","能力结构","来源","是否有视频","视频文件","视频质量","视频类型"
                ,"第三级知识点1","第三级知识点2","第三级知识点3","第三级知识点4","第三级知识点5"};
        List<String> headList = Arrays.asList(head);
        Row r0 = s.createRow(0);
        for(int cellnum = 0; cellnum < head.length; cellnum ++) {
            Cell c = r0.createCell(cellnum);
            c.setCellValue(head[cellnum]);
            c.setCellStyle(style);
        }


        for(int rownum = 1; rownum <= list.size(); rownum++) {
            Row r = s.createRow(rownum);
            Timu timu = list.get(rownum-1);
            Cell cTiGan = r.createCell(headList.indexOf("题干"));
            cTiGan.setCellValue(timu.getTigan());

            Cell cXuanXiang = r.createCell(headList.indexOf("备选答案"));
            cXuanXiang.setCellValue(timu.getXuanxiang());

            Cell cDaAn = r.createCell(headList.indexOf("正确答案"));
            cDaAn.setCellValue(timu.getDaan());

            Cell cJieXi = r.createCell(headList.indexOf("解析"));
            cJieXi.setCellValue(timu.getJiexi());

            Cell cTiXing = r.createCell(headList.indexOf("题型*"));
            cTiXing.setCellValue(timu.getTixing());

            String[] zsdArr = timu.getZsdArr();
            if (null != zsdArr && zsdArr.length > 0) {
                for (int i = 1; i <= zsdArr.length; i++) {
                    Cell cZsd = r.createCell(headList.indexOf("第三级知识点"+i));
                    cZsd.setCellValue(zsdArr[i-1]);
                }
            }
        }
        String filename = "D:\\workbook.xls";
        if(wb instanceof XSSFWorkbook) {
            filename = filename + "x";
        }
        FileOutputStream out = new FileOutputStream(filename);
        wb.write(out);
        out.close();
    }
}
