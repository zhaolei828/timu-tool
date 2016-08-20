package com.gun.tm.tool.excel;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

/**
 * @author zhaolei
 * @create 2016-08-16 11:19
 */
public class Main4 {
    public static void main(String[] args) throws IOException {
        final String path = "F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例2\\湖北襄阳本地化·七年级数学·第十章试题6.13\\";
        final String file = "2016年湖北省襄阳市新人教版七年级数学下学期第十章《数据的收集、整理与描述》测试题数学试卷答案解析试卷编号：04016230207008.doc";
        InputStream input = new FileInputStream(path + file);
        HWPFDocument wordDocument = new HWPFDocument(input);
        Range range = wordDocument.getRange();
        int size = range.numParagraphs();
        for (int j = 0; j < size; j++) {
            Paragraph paragraph = range.getParagraph(j);
            System.out.println(j + "行：" + paragraph.text().trim());
        }
//        String[] s = new String[50];
//        TableIterator it = new TableIterator(range);
//        int index = 0;
//        while (it.hasNext()) {
//            Table tb = (Table) it.next();
//            for (int i = 0; i < tb.numRows(); i++) {
//                //System.out.println("Numrows :"+tb.numRows());
//                TableRow tr = tb.getRow(i);
//                for (int j = 0; j < tr.numCells(); j++) {
//                    //System.out.println("numCells :"+tr.numCells());
////                      System.out.println("j   :"+j);
//                    TableCell td = tr.getCell(j);
//                    for (int k = 0; k < td.numParagraphs(); k++) {
//                        //System.out.println("numParagraphs :"+td.numParagraphs());
//                        Paragraph para = td.getParagraph(k);
//                        s[index] = para.text().trim();
//                        index++;
//                    }
//                }
//            }
//        }
//        for(int i=0;i<s.length;i++){
//            System.out.println(s[i]);
//        }
    }
    public static void poiWordTableReplace(String sourceFile, String newFile,
                                           Map<String, Text> replaces) throws Exception {
        FileInputStream in = new FileInputStream(sourceFile);
        HWPFDocument hwpf = new HWPFDocument(in);
        Range range = hwpf.getRange();// 得到文档的读取范围
        TableIterator it = new TableIterator(range);
        // 迭代文档中的表格
        while (it.hasNext()) {
            Table tb = (Table) it.next();
            // 迭代行，默认从0开始
            for (int i = 0; i < tb.numRows(); i++) {
                TableRow tr = tb.getRow(i);
                // 迭代列，默认从0开始
                for (int j = 0; j < tr.numCells(); j++) {
                    TableCell td = tr.getCell(j);// 取得单元格
                    // 取得单元格的内容
                    for (int k = 0; k < td.numParagraphs(); k++) {
                        Paragraph para = td.getParagraph(k);

                        String s = para.text();
                        final String old = s;
                        for (String key : replaces.keySet()) {
                            if (s.contains(key)) {
                                s = s.replace(key, replaces.get(key).getText());
                            }
                        }
                        if (!old.equals(s)) {// 有变化
                            para.replaceText(old, s);
                            s = para.text();
                            System.out.println("old:" + old + "->" + "s:" + s);
                        }

                    } // end for
                } // end for
            } // end for
        } // end while

        FileOutputStream out = new FileOutputStream(newFile);
        hwpf.write(out);

        out.flush();
        out.close();

    }
    public abstract class Text {

        public abstract String getText();

        public Text str(final String string) {
            return new Text() {
                @Override
                public String getText() {
                    return string;
                }
            };
        }

    }
}
