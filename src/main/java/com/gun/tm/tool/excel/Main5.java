package com.gun.tm.tool.excel;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author zhaolei
 * @create 2016-08-20 14:03
 */
public class Main5 {
    public static void main(String[] args) throws IOException {
        final String path = "F:\\项目\\其他\\转excel\\word导入excel\\word导入excel\\样例\\样例2\\湖北襄阳本地化·七年级数学·第十章试题6.13\\";
        final String file = "2016年湖北省襄阳市新人教版七年级数学下学期第十章《数据的收集、整理与描述》测试题数学试卷答案解析试卷编号：04016230207008.doc";
        InputStream is = new FileInputStream(path+file);
        HWPFDocument doc = new HWPFDocument(is);
        //输出文本
        System.out.println(doc.getDocumentText());
        Range range = doc.getRange();
//    this.insertInfo(range);
        printInfo(range);
        //读表格
        readTable(range);
        //读列表
        readList(range);
        //删除range
        //Range r = new Range(2, 5, doc);
        //r.delete();//在内存中进行删除，如果需要保存到文件中需要再把它写回文件
        //把当前HWPFDocument写到输出流中
       // doc.write(new FileOutputStream("D:\\test.doc"));
        closeStream(is);
    }
    private static void readTable(Range range) {
        //遍历range范围内的table。
        TableIterator tableIter = new TableIterator(range);
        Table table;
        TableRow row;
        TableCell cell;
        while (tableIter.hasNext()) {
            table = tableIter.next();
            int rowNum = table.numRows();
            for (int j=0; j<rowNum; j++) {
                row = table.getRow(j);
                int cellNum = row.numCells();
                for (int k=0; k<cellNum; k++) {
                    cell = row.getCell(k);
                    //输出单元格的文本
                    System.out.println(cell.text().trim());
                }
            }
        }
    }
    private static void closeStream(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    private static void readList(Range range) {
        int num = range.numParagraphs();
        Paragraph para;
        for (int i=0; i<num; i++) {
            para = range.getParagraph(i);
            if (para.isInList()) {
                System.out.println("list: " + para.text());
            }
        }
    }
    private static void printInfo(Range range) {
        //获取段落数
        int paraNum = range.numParagraphs();
        System.out.println(paraNum);
        int secNum = range.numSections();
        System.out.println(secNum);
        Section section;
        for (int i=0; i<secNum; i++) {
            section = range.getSection(i);
            System.out.println(section.getMarginLeft());
            System.out.println(section.getMarginRight());
            System.out.println(section.getMarginTop());
            System.out.println(section.getMarginBottom());
            System.out.println(section.getPageHeight());
            System.out.println(section.text());
        }
    }
}
