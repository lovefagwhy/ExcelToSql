package com.excel;


import java.io.File;
import java.util.Map;

public class ExcelToSql {

    public static void main(String[] args) {
        try {
            //获取录入日期目录
            String dirName = FileUtil.getFileDir();
//            String dirName = "20201010";
            if (dirName == null) {
                System.out.println("输入为空，结束程序");
                return;
            }
            String path = FileUtil.getPath();
            System.out.println("项目所在路径：" + path);
            File file = new File(path + "/InportExcel/" + dirName);
            if (file.isDirectory()) {
                System.out.println("                                                  ");
                System.out.println("进入目录:" + dirName);
                File[] files = file.listFiles();
                if (files == null || files.length == 0) {
                    System.out.println("目录中没有文件");
                    return;
                }
                int aCount = 0;
                int bCount = 0;
                int cCount = 0;
                Map<String, String> configProps = PropertyUtil.getConfigProps(path);
                for (File f : files) {
                    String fileName = f.getName();
                    System.out.println("                                                  ");
                    if (fileName.startsWith("A") && (fileName.endsWith(".xlsx")||fileName.endsWith(".xls"))) {
                        aCount++;
                        if (aCount > 1) {
                            System.out.println(fileName + "不是第一个带A开头Excel表格，只解析第一个");
                            FileUtil.moveFile(dirName, path, f, false);
                            continue;
                        }
                        System.out.println("开始解析表格:" + fileName);
                        Boolean aSuc = PoiUtil.parseAExcel(f, path, dirName,configProps);
                        FileUtil.moveFile(dirName, path, f, aSuc);
                    } else if (fileName.startsWith("B")&& (fileName.endsWith(".xlsx")||fileName.endsWith(".xls"))) {
                        System.out.println("                                                  ");
                        bCount++;
                        if (bCount > 1) {
                            System.out.println(fileName + "不是第一个带B开头Excel表格，只解析第一个");
                            FileUtil.moveFile(dirName, path, f, false);
                            continue;
                        }
                        System.out.println("开始解析表格:" + fileName);
                        Boolean bSuc = PoiUtil.parseBExcel(f, path, dirName,configProps);
                        FileUtil.moveFile(dirName, path, f, bSuc);
                    } else if (fileName.startsWith("C")&& (fileName.endsWith(".xlsx")||fileName.endsWith(".xls"))) {
                        System.out.println("                                                  ");
                        cCount++;
                        if (cCount > 1) {
                            System.out.println(fileName + "不是第一个带C开头Excel表格，只解析第一个");
                            FileUtil.moveFile(dirName, path, f, false);
                            continue;
                        }
                        System.out.println("开始解析表格:" + fileName);
                        Boolean cSuc = PoiUtil.parseCExcel(f, path, dirName,configProps);
                        FileUtil.moveFile(dirName, path, f, cSuc);
                    } else {
                        System.out.println("                                                  ");
                        FileUtil.moveFile(dirName, path, f, false);
                    }
                }
                files = file.listFiles();
                if (files == null || files.length == 0) {
                    file.delete();
                }
            } else {
                System.out.println("                                                  ");
                System.out.println(file.getName() + "不是目录");
                return;
            }

        } catch (Exception e) {
            System.out.println("                                                  ");
            e.printStackTrace();
            System.out.print("出错");
        }
    }


}
