package com.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

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
//        String string = "辽\n33,23宁/(12)!d..。，;；；】、大@#$%^&*~\r\n天津（1）";
//        String s = string.replaceAll("(\r|\n|[0-9]|[a-zA-Z]|（|）|\\.|\\,|\\，|\\。|/|\\(|\\)|#|\\^|@|\\$|%|&|\\*|!|\\~)*", "");
//        System.out.println(s);
//        String s1 = string.replaceAll("[^\\u4e00-\\u9fa5]", "");
//        System.out.println(s1);
        List<Map<String,String>> orders = new ArrayList<>();
        for (int j = 1; j < 31; j++) {
            Map<String,String> map = new HashMap<>();
            for (int i = 1; i <5 ; i++) {
                map.put(i+"",j+"");
            }
            orders.add(map);
        }
        dealExcel(orders);
    }

    public static void dealExcel(List<Map<String,String>> orders) {
        HSSFWorkbook wb = new HSSFWorkbook();
        //sheet的名字
        HSSFSheet sheet = wb.createSheet("订单列表");
        //设置默认列宽和行高
        sheet.setDefaultColumnWidth((short) 20);
        sheet.setDefaultRowHeightInPoints(20);
        //标题行，第0行数据,就是图片上的蓝色行
        HSSFRow row = sheet.createRow(0);
        String[] str = {"业务员", "客户姓名", "客户电话","備註"};
//        第一行
        OutputStream out = null;
        try {
            //设置样式
            HSSFCellStyle cellStyle = wb.createCellStyle();
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(HSSFColor.ROYAL_BLUE.index);
            //第一行数据
            for (int i = 0; i < str.length; i++) {
                //填充背景颜色
                row.createCell(i).setCellStyle(cellStyle);
                row.getCell(i).setCellValue(str[i]);
            }
            //拿出所有的数据
            Iterator<Map<String,String>> iterator = orders.iterator();
            //一条数据就是一行记录,开始起始的行号
            int index = 1;
            while (iterator.hasNext()) {
                //这个作用是为了记录需要合并的行的起始行
                int temp = index;
                Map<String,String> goods = iterator.next();
                int num = 0;
                if (null != goods && goods.size() > 0) {
                    num += goods.size();
                }
                //合并行的结束行号
                index += num;
                //以不需要合并行的个数，作为需要创建的行数
                for (int e = 0; e < num; e++) {
                    HSSFRow sheetRow = sheet.createRow(temp + e);
                    for (int i = 0; i < 4; i++) {
                        //填充不需要合并列的数据信息  start
                        if (i == 0) {
                            sheetRow.createCell(i).setCellValue(goods.get(i+""));
                            continue;
                        }
                        if (i == 1) {
                            sheetRow.createCell(i).setCellValue(goods.get(i+""));
                            continue;
                        }
                        if (i == 2) {
                            sheetRow.createCell(i).setCellValue(goods.get(i+""));
                            continue;
                        }
                        //填充需要合并行的单元格的信息
                        sheetRow.createCell(i).setCellValue("beizhu");
                    }
                    //合并单元格
//                    if (e == num - 1) {
//                        /**
//                         23代表的意思是并不需要合并的列的   起始列号+ 1 ;
//                         比如：我的第 6列需要合并列的起始位置，则此处写  7
//                         */
//                        for (int j = 0; j < 3; j++) {
//                            //行号   行号  列号  列号
//                            sheet.addMergedRegion(new CellRangeAddress(temp, index - 1, j, j));
//                        }
//                    }
                }
                //输出到本地
                OutputStream o = null;
                try {
                    o = new FileOutputStream("D://2007.xls");
                    wb.write(o);
                    o.close();
                } catch (Exception e) {
                    e.printStackTrace();
                } finally {
                    if (o != null) {
                        o.close();
                    }
                    if (out != null) {
                        out.close();
                    }
                }
            }
        } catch (Exception e1) {
            e1.printStackTrace();
        }
    }
}
