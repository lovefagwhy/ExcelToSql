package com.excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.List;
import java.util.Scanner;

/**
 * Description:
 * Author:        liuzhuang
 * Create Date:   2020/4/9 10:04
 */
public class FileUtil {
    //移动到指定目录
    public static void moveFile(String dirName, String path, File f, Boolean aSuc) {
        String name = f.getName();
        if (aSuc) {
            File file = new File(path + "/HistoryExcel/" + dirName);
            if (!file.exists()) {
                file.mkdir();
            }
            File hFile = new File(path + "/HistoryExcel/" + dirName + "/" + name);
            if(hFile.exists()){
                hFile.delete();
            }
            f.renameTo(hFile);
            System.out.println(name + "移动到成功目录");
        } else {
            File file = new File(path + "/ErrorExcel/" + dirName);
            if (!file.exists()) {
                file.mkdir();
            }
            File eFile = new File(path + "/ErrorExcel/" + dirName + "/" + name);
            if(eFile.exists()){
                eFile.delete();
            }
            f.renameTo(eFile);
            System.out.println(name + "移动到失败目录");
        }
    }

    public static String getFileDir() {
        System.out.println("请输入导入EXCEL所在目录名称，格式[yyyyMMdd,例如:20200101]：");
        Scanner scanner = null;
        String str = null;
        try {
            scanner = new Scanner(System.in);
            if (scanner.hasNextLine()) {
                str = scanner.nextLine();
                if (str == null || "".equals(str.trim())) {
                    return null;
                }
                System.out.println("输入的数据为：" + str);
            }
            scanner.close();
        } catch (Exception e) {
            if (scanner != null) {
                scanner.close();
            }
        } finally {
            if (scanner != null) {
                scanner.close();
            }
        }
        return str;
    }

    public static String getPath() {
        String path = PoiUtil.class.getProtectionDomain().getCodeSource().getLocation().getPath();
//        String path = "D:/DATACompany/IdeaWorkSpace/ExcelToSql/out/artifacts/ExcelToSql_jar";
        if (path.startsWith("/")) {
            path = path.substring(1);
        }
        if (path.contains(".jar")) {
            path = path.substring(0, path.lastIndexOf("/"));
        }
        System.out.println("                                                  ");
        return path;
    }

    /**
     * Method description :
     * datas:sql集合，第一条记录为title，sql从第二条开始
     * path：项目路径
     * dirName：录入参数(日期目录)
     * fName：xls文件名
     */
    public static boolean makeSql(List<String> datas, String path, String dirName, String fName) throws Exception {
        BufferedWriter bw = null;
        try {
            File file = new File(path + "/ExportSql/" + dirName);
            if (!file.exists()) {
                file.mkdir();
            }
            fName = fName.substring(0, fName.lastIndexOf("."));
            bw = new BufferedWriter(new FileWriter(path + "/ExportSql/" + dirName + "/" + fName + dirName + ".sql"));

            for (String data : datas) {
                bw.write(data);
            }
            //刷新流
            bw.flush();

            //关闭资源
            bw.close();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("写sql文件出错");
            if (bw != null) {
                bw.close();
            }
            return false;
        } finally {
            if (bw != null) {
                try {
                    bw.close();
                } catch (Exception e) {
                    return false;
                }
            }
            return true;
        }
    }
}
