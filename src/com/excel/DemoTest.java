package com.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

/**
 * Description:
 * Author:        liuzhuang
 * Create Date:   2020/4/15 14:16
 */
public class DemoTest {
    public static void main(String[] args) {
        //测试字符串通过正则 过滤非中文字符
//        String string = "辽\n33,23宁/(12)!d..。，;；；】、大@#$%^&*~\r\n天津（1）";
//        String s = string.replaceAll("(\r|\n|[0-9]|[a-zA-Z]|（|）|\\.|\\,|\\，|\\。|/|\\(|\\)|#|\\^|@|\\$|%|&|\\*|!|\\~)*", "");
//        System.out.println(s);
//        String s1 = string.replaceAll("[^\\u4e00-\\u9fa5]", "");
//        System.out.println(s1);
        //打印文件目录下所有文件名称
        printFileName();
    }

    public static void printFileName() {
        File file = new File("d:/项目文件/与国税交互项目");
        File[] files = file.listFiles();
        for (File dir : files) {
            if(dir.isDirectory()){
                System.out.println("项目名称："+dir.getName());
                showFileName(dir);

            }
        }
    }
    /**
     * @param f
     * @param size 用来控制空格个数
     */
    public static void showFileName(File f) {
        if(f.isDirectory()){
            File[] files = f.listFiles();
            for (int i = 0; i < files.length; i++) {
                if(files[i].isDirectory()){
                    if (files[i].listFiles().length != 0) {
                        showFileName(files[i]);
                    }
                }else {
                    System.out.println(files[i].getName());
                }
            }
        }else{
            System.out.println(f.getName());
        }

    }

}
