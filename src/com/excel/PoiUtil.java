package com.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * @Author: LiuZhuang
 * Description:
 * Date:Created in 21:20 2020/4/8.
 * Modified By:
 */
public class PoiUtil {
    public static StringBuilder SQL_A_PRE = new StringBuilder("INSERT INTO SW_BORDER_INFO(AREA,PORT,BORDER_PORT,PORT_TYPE,BORDER_COUNTRY,P_STATUS,IN_PERSON,In_Passport,OUT_PERSON,Out_Passport,G_STATUS,IN_GOODS,OUT_GOODS,NOTE,BORDER_TYPE,START_DATE,END_DATE) VALUES (");
    public static StringBuilder SQL_B_PRE = new StringBuilder("INSERT INTO SW_BORDER_INFO(AREA,PORT,BORDER_PORT,PORT_TYPE,BORDER_COUNTRY,G_STATUS,P_STATUS,BORDER_TYPE,IN_PERSON,OUT_PERSON,IN_PASSPORT,IN_PASSPORT_FOREIGN,OUT_PASSPORT,OUT_PASSPORT_FOREIGN,In_DRIVERS,Out_DRIVERS,NOTE,START_DATE,END_DATE) VALUES(");
    public static StringBuilder SQL_C_PRE = new StringBuilder("INSERT INTO SW_BORDER_INFO(AREA,PORT,PORT_TYPE,G_STATUS,BORDER_TYPE,IN_PERSON,OUT_PERSON,In_DRIVERS,Out_DRIVERS,NOTE,START_DATE,END_DATE) VALUES(");
    public static String SQL_SUF = ");\r\n";
    public final static String XLS = "xls";

    //边区
    public static Boolean parseAExcel(File f, String path, String dirName, Map<String, String> props,Map<String,String> portProps) throws Exception {
        FileInputStream fi = null;
        Workbook wb;
        boolean isTr = false;
        try{
            fi = new FileInputStream(f);
            // 根据文件格式(2003或者2007)来初始化
            if (f.getName().endsWith("xlsx")) {
                wb = new XSSFWorkbook(fi);
            } else {
                wb = new HSSFWorkbook(fi);
            }
            Sheet sheet = wb.getSheetAt(0);
            if (sheet == null) {
                return false;
            }
            //获取 口岸信息
            List<String> datas = new ArrayList<>();
            //获取第一行文本内容  标题
            String row0 = sheet.getRow(0).getCell(0).getStringCellValue();
            //获取第二行文本内容
            String row1 = sheet.getRow(1).getCell(0).getStringCellValue();
            if(row0!=null && row0.contains("边境口岸运行")){
                datas.add("--"+row0+"\r\n");
            }else if(row1!=null && row1.contains("边境口岸运行")){
                datas.add("--"+row1+"\r\n");
            }
            //拆分合并单元格
            removeMerge(sheet, props.get("removeMergeA"));
            //获取解析 坐标点
            int start = getStart(sheet);
            String subNumA = props.get("subNumA");
            int num;
            try {
                num = Integer.parseInt(subNumA);
            } catch (Exception e) {
                num = 77;
            }
            //中文字段  去除非中文字符正则
            String reg = props.get("REG");
            if(reg == null ||"".equals(reg)){
                reg = "[^\\u4e00-\\u9fa5]";
            }
            //是否需要将口岸性质 里的水/陆运  改成水运
            String isChange = props.get("IS_PORT_TYPE_CHANGE");
            if(isChange == null ||"".equals(isChange)){
                isChange = "0";
            }
            int end = start+num;
            // 循环行Row 从第六行开始
            for (int rowNum = start; rowNum < end; rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
                // 循环列Cell
                StringBuilder tempSql=new StringBuilder(SQL_A_PRE);
                for (int cellNum = 1; cellNum <= row.getLastCellNum(); cellNum++) {
                    Cell cell = row.getCell(cellNum);
                    if(cell !=null){
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                    }
                    switch(cellNum){
                        case 1: // 省份（B列） AREA 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacernNum(cell.getStringCellValue(),reg)+"',");
                            break;
                        case 2: //口岸名称（C列） PORT 字符型  根据口岸名称（C列）查找对应的国外口岸 BORDER_PORT 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacernNum(cell.getStringCellValue(),reg)+"',");
                            if(isNull(cell)){
                                tempSql.append(null+",");
                            }else{
                                String border_port = portProps.get(replacernNum(cell.getStringCellValue(),reg));
                                tempSql.append(isNullStr(border_port)?null+",":"'"+border_port+"',");
                            }
                            break;
                        case 3: //  性质（D列） PORT_TYPE 字符型
                            if(isNull(cell)){
                                tempSql.append(null+",");
                            }else{
                                String port_type = replacern(cell.getStringCellValue());
                                if(!"0".equals(isChange)){
                                    if(port_type.contains("水/陆")){
                                        port_type = "公路";
                                    }
                                }
                                tempSql.append(isNullStr(port_type)?null+",":"'"+port_type+"',");
                            }
                            break;
                        case 4: //国家（E列） BORDER_COUNTRY 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacernNum(cell.getStringCellValue(),reg)+"',");
                            break;
                        case 5: //客运状态（F列） P_STATUS 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacern(cell.getStringCellValue())+"',");
                            break;
                        case 6://入境人次（总数）（G列） IN_PERSON 字符型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 7: //入境人次（护照）（H列）In_Passport 数字型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 8: //出境人次（总数）（I列） OUT_PERSON 字符型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 9: //出境人次（护照数）（J列） Out_Passport 数字型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 10: //货运状态(K列)  G_STATUS 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacern(cell.getStringCellValue())+"',");
                            break;
                        case 11: //进口货物(吨)（L列） IN_GOODS 字符型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 12: //出口货物(吨)（M列） OUG_GOODS 字符型
                            tempSql.append(isNull(cell)?"0,":"'"+replacernd(cell.getStringCellValue())+"',");
                            break;
                        case 13: //备注（N列） NOTE 字符型
                            tempSql.append(isNull(cell)?null+",":"'"+replacern(cell.getStringCellValue())+"',");
                            break;
                        default:
                            break;
                    }
                }
                tempSql.append("1,");
                tempSql.append("to_date('"+dirName+"','yyyyMMdd'),");
                tempSql.append("to_date('"+dirName+"','yyyyMMdd')");
                tempSql.append(SQL_SUF);
                datas.add(tempSql.toString());
                System.out.println(tempSql.toString());
                isTr = FileUtil.makeSql(datas,path,dirName,f.getName());
            }
            //关闭输入流
            fi.close();
            return isTr;
        } catch (Exception e){
            e.printStackTrace();
            System.out.println("解析Excel失败");
            return false;
        }finally {
            try{
                if(fi!=null){
                    fi.close();
                }
            }catch (Exception e){
                return false;
            }
            return  isTr;
        }
    }

    /**
     * Method description :
     * 广东毗邻港澳口岸运行状况一览表拆分
     * f：表格文件
     * path：项目路径
     * dirName:表格所在目录，为cmd窗口录入
     * Author：  liuzhuang
     */
    public static Boolean parseBExcel(File f, String path, String dirName, Map<String, String> props,Map<String,String> portProps) {
        //创建工作簿对象
        Workbook wb;
        FileInputStream fi = null;
        boolean suc = false;
        try {
            fi = new FileInputStream(f);
            // 判断是否是excel2007格式
            boolean isE2007 = false;
            if (f.getName().endsWith("xlsx")) {
                isE2007 = true;
            }

            // 根据文件格式(2003或者2007)来初始化
            if (isE2007) {
                wb = new XSSFWorkbook(fi);
            } else {
                wb = new HSSFWorkbook(fi);
            }

            Sheet sheet = wb.getSheetAt(0);
            //先返回XSSF和HSSF对象，再创建一个用于计算公式单元格的对象
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            /*双for循环遍历工作簿中单元格*/
            List<String> datas = new ArrayList<>();
            //拆分合并单元格
            removeMerge(sheet, props.get("removeMergeB"));
            Map<String, String> strMap;
            Map<String, Double> numMap;
            StringBuilder tempSql;
            String subNumB = props.get("subNumB");
            int num;
            int start = 5;
            try {
                num = Integer.parseInt(subNumB);
            } catch (Exception e) {
                num = 12;
            }
            //中文字段  去除非中文字符正则
            String reg = props.get("REG");
            if(reg == null ||"".equals(reg)){
                reg = "[^\\u4e00-\\u9fa5]";
            }
            //是否需要将口岸性质 里的水/陆运  改成水运
            String isChange = props.get("IS_PORT_TYPE_CHANGE");
            if(isChange == null ||"".equals(isChange)){
                isChange = "0";
            }

            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                //指定单元格
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(Short.parseShort(0 + ""));
                if(cell.getCellType() == 0){
                    double numericCellValue = cell.getNumericCellValue();
                    if(numericCellValue == 1.0){
                        start = i;
                        break;
                    }

                }else if(cell.getCellType() == 1){
                    String index = cell.getStringCellValue();
                    String title = sheet.getRow(i).getCell(0).getStringCellValue();
                    if(title!=null && title.contains("广东毗邻港")){
                        datas.add("--" + title + "\r\n");
                        continue;
                    }
                    if("1".equals(index)){
                        start = i;
                        break;
                    }
                }
            }
            //行循环
            for (int i = start; i < start+num; i++) {
                tempSql = new StringBuilder(SQL_B_PRE);
                strMap = new HashMap<>();
                numMap = new HashMap<>();
                //行对象
                Row row = sheet.getRow(i);
                if(row == null){
                    continue;
                }
                //取最后一列列号
                int cellNum = row.getLastCellNum();
                //列循环
                for (int j = 1; j < cellNum; j++) {
                    //指定单元格
                    Cell cell = row.getCell(Short.parseShort(j + ""));
                    //单元格值对象
                    CellValue c = evaluator.evaluate(cell);
                    //判断单元格是否有值
                    if (c != null) {
                        switch (c.getCellType()) {
                            case 1:
                                //得到单元格值
                                String value = c.getStringValue();
                                switch (j) {
                                    case 1:
                                        strMap.put("AREA", replacernNum(value,reg));
                                        break;
                                    case 2:
                                        strMap.put("NOTE", value.replaceAll("\r|\n*",""));
                                        break;
                                    case 3:
                                        strMap.put("PORT", replacernNum(value,reg));
                                        strMap.put("AREA", "广东");
                                        break;
                                    case 4:
                                        strMap.put("PORT_TYPE", replacern(value));
                                        break;
                                    case 5:
                                        strMap.put("BORDER_COUNTRY", replacernNum(value,reg));
                                        break;
                                    case 6:
                                        strMap.put("P_STATUS", replacern(value));
                                        break;
                                    case 13:
                                        strMap.put("G_STATUS", replacern(value));
                                        break;
                                    case 7:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("IN_PRESON_1", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("IN_DRIVERS", numMap.get("IN_PRESON_1"));
                                        break;
                                    case 8:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("IN_PRESON_2", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("IN_PASSPORT_1", numMap.get("IN_PRESON_2"));
                                        break;
                                    case 9:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("IN_PRESON_3", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("IN_PASSPORT_2", numMap.get("IN_PRESON_3"));
                                        break;
                                    case 10:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("OUT_PRESON_1", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("OUT_DRIVERS", numMap.get("OUT_PRESON_1"));
                                        break;
                                    case 11:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("OUT_PRESON_2", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("OUT_PASSPORT_1", numMap.get("OUT_PRESON_2"));
                                        break;
                                    case 12:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("OUT_PRESON_3", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("OUT_PASSPORT_2", numMap.get("OUT_PRESON_3"));
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            case 0:
                                //得到单元格内数字
                                Double dvalue = c.getNumberValue();
                                switch (j) {
                                    case 1:
                                        strMap.put("AREA", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 2:
                                        strMap.put("NOTE", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 3:
                                        strMap.put("PORT", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 4:
                                        strMap.put("PORT_TYPE", dvalue+"");
                                        break;
                                    case 5:
                                        strMap.put("BORDER_COUNTRY", dvalue+"");
                                        break;
                                    case 6:
                                        strMap.put("P_STATUS", dvalue+"");
                                        break;
                                    case 13:
                                        strMap.put("G_STATUS", dvalue+"");
                                        break;
                                    case 7:
                                        numMap.put("IN_PRESON_1", dvalue);
                                        numMap.put("IN_DRIVERS", dvalue);
                                        break;
                                    case 8:
                                        numMap.put("IN_PRESON_2", dvalue);
                                        numMap.put("IN_PASSPORT_1", dvalue);
                                        break;
                                    case 9:
                                        numMap.put("IN_PRESON_3", dvalue);
                                        numMap.put("IN_PASSPORT_2", dvalue);
                                        break;
                                    case 10:
                                        numMap.put("OUT_PRESON_1", dvalue);
                                        numMap.put("OUT_DRIVERS", dvalue);
                                        break;
                                    case 11:
                                        numMap.put("OUT_PRESON_2", dvalue);
                                        numMap.put("OUT_PASSPORT_1", dvalue);
                                        break;
                                    case 12:
                                        numMap.put("OUT_PRESON_3", dvalue);
                                        numMap.put("OUT_PASSPORT_2", dvalue);
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            default:
                        }
                    }
                }
                String area = strMap.get("AREA");
                tempSql.append((area == null || "".equals(area)) ? null + "," : "'" + area + "',");
                String port = strMap.get("PORT");
                tempSql.append((port == null || "".equals(port)) ? null + "," : "'" + port + "',");
                if(port!=null && !"".equals(port)){
                    String border_port = portProps.get(port);
                    tempSql.append((border_port == null || "".equals(border_port)) ? null + "," : "'" + border_port + "',");
                }else{
                    tempSql.append(null + ",");
                }
                String port_type = strMap.get("PORT_TYPE")== null ? null :strMap.get("PORT_TYPE");
                if(port_type == null){
                    tempSql.append(null+",");
                }else if(port_type.startsWith("水") && "1".equals(isChange)){
                    tempSql.append("'水运',");
                }else{
                    tempSql.append("'"+port_type+"',");
                }
                String border_country = strMap.get("BORDER_COUNTRY");
                tempSql.append(border_country == null ? null + "," : "'" + border_country + "',");
                String gStatus = strMap.get("G_STATUS");
                tempSql.append(gStatus == null ? null + "," : "'" + gStatus + "',");
                String pStatus = strMap.get("P_STATUS");
                tempSql.append(pStatus == null ? null + "," : "'" + pStatus + "',");
                tempSql.append("'2',");
                tempSql.append((numMap.get("IN_PRESON_1") == null ? 0 : numMap.get("IN_PRESON_1").intValue()) + (numMap.get("IN_PRESON_2") == null ? 0 : numMap.get("IN_PRESON_2").intValue()) + (numMap.get("IN_PRESON_3") == null ? 0 : numMap.get("IN_PRESON_3").intValue()));
                tempSql.append(",");
                tempSql.append((numMap.get("OUT_PRESON_1") == null ? 0 : numMap.get("OUT_PRESON_1").intValue()) + (numMap.get("OUT_PRESON_2") == null ? 0 : numMap.get("OUT_PRESON_2").intValue()) + (numMap.get("OUT_PRESON_3") == null ? 0 : numMap.get("OUT_PRESON_3").intValue()));
                tempSql.append(",");
                //入境护照总数
                tempSql.append((numMap.get("IN_PASSPORT_1") == null ? 0 : numMap.get("IN_PASSPORT_1").intValue()) + (numMap.get("IN_PASSPORT_2") == null ? 0 : numMap.get("IN_PASSPORT_2").intValue()));
                tempSql.append(",");
                //外方入境护照数量
                tempSql.append(numMap.get("IN_PASSPORT_2") == null ? 0 : numMap.get("IN_PASSPORT_2").intValue());
                tempSql.append(",");
                //出境护照总数
                tempSql.append((numMap.get("OUT_PASSPORT_1") == null ? 0 : numMap.get("OUT_PASSPORT_1").intValue()) + (numMap.get("OUT_PASSPORT_2") == null ? 0 : numMap.get("OUT_PASSPORT_2").intValue()));
                tempSql.append(",");
                //外方出境护照数量
                tempSql.append(numMap.get("OUT_PASSPORT_2") == null ? 0 : numMap.get("OUT_PASSPORT_2").intValue());
                tempSql.append(",");
                tempSql.append(numMap.get("IN_DRIVERS") == null ? 0 : numMap.get("IN_DRIVERS").intValue());
                tempSql.append(",");
                tempSql.append(numMap.get("OUT_DRIVERS") == null ? 0 : numMap.get("OUT_DRIVERS").intValue());
                tempSql.append(",");
                String note = strMap.get("NOTE");
                tempSql.append((note == null || "".equals(note)) ? null + "," : "'" + note + "',");
                tempSql.append("to_date('" + dirName + "','yyyyMMdd'),");
                tempSql.append("to_date('" + dirName + "','yyyyMMdd')");
                tempSql.append(SQL_SUF);
                datas.add(tempSql.toString());
                System.out.println(tempSql.toString());
            }
            suc = FileUtil.makeSql(datas, path, dirName, f.getName());
            //关闭输入流
            fi.close();
        } catch (Exception e){
            e.printStackTrace();
            System.out.println("解析Excel失败");
            suc = false;
        }finally {
            try{
                if(fi!=null){
                    fi.close();
                }
            }catch (Exception e){
                suc = false;
            }
            return  suc;
        }
    }


    /**
     * Method description :
     * 水运表拆分
     * f：表格文件
     * path：项目路径
     * dirName:表格所在目录，为cmd窗口录入
     * Author：  liuzhuang
     */
    public static Boolean parseCExcel(File f, String path, String dirName, Map<String, String> props) {
        //创建工作簿对象
        Workbook wb;
        FileInputStream fi = null;
        boolean suc = false;
        try {
            fi = new FileInputStream(f);
            // 判断是否是excel2007格式
            boolean isE2007 = false;
            if (f.getName().endsWith("xlsx")) {
                isE2007 = true;
            }

            // 根据文件格式(2003或者2007)来初始化
            if (isE2007) {
                wb = new XSSFWorkbook(fi);
            } else {
                wb = new HSSFWorkbook(fi);
            }

            Sheet sheet = wb.getSheetAt(0);
            //先返回XSSF和HSSF对象，再创建一个用于计算公式单元格的对象
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            /*双for循环遍历工作簿中单元格*/
            List<String> datas = new ArrayList<>();

            //拆分合并单元格
            removeMerge(sheet, props.get("removeMergeC"));
            Map<String, String> strMap;
            Map<String, Double> numMap;
            StringBuilder tempSql;
            //行循环
            String subNumC = props.get("subNumC");
            int num;
            int start = 3;
            try {
                num = Integer.parseInt(subNumC);
            } catch (Exception e) {
                num = 125;
            }
            //中文字段  去除非中文字符正则
            String reg = props.get("REG");
            if(reg == null ||"".equals(reg)){
                reg = "[^\\u4e00-\\u9fa5]";
            }
            //获取序号为1的行索引
            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                //指定单元格
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(Short.parseShort(0 + ""));
                if(cell.getCellType() == 0){
                    double numericCellValue = cell.getNumericCellValue();
                    if(numericCellValue == 1.0){
                        start = i;
                        break;
                    }

                }else if(cell.getCellType() == 1){
                    String index = cell.getStringCellValue();
                    String title = sheet.getRow(i).getCell(0).getStringCellValue();
                    if(title!=null && title.contains("全国水运口岸")){
                        datas.add("--" + title + "\r\n");
                        continue;
                    }
                    if("1".equals(index)){
                        start = i;
                        break;
                    }
                }
            }
            //行循环
            for (int i = start; i < start+num; i++) {
                tempSql = new StringBuilder(SQL_C_PRE);
                strMap = new HashMap<>();
                numMap = new HashMap<>();
                //行对象
                Row row = sheet.getRow(i);
                if(row == null){
                    continue;
                }
                //取最后一列列号
                int cellNum = row.getLastCellNum();
                //列循环
                for (int j = 1; j < cellNum; j++) {
                    //指定单元格
                    Cell cell = row.getCell(Short.parseShort(j + ""));
                    //单元格值对象
                    CellValue c = evaluator.evaluate(cell);
                    //判断单元格是否有值
                    if (c != null) {
                        switch (c.getCellType()) {
                            case 1:
                                //得到单元格值
                                String value = c.getStringValue();
                                switch (j) {
                                    case 1:
                                        strMap.put("AREA", replacernNum(value,reg));
                                        break;
                                    case 2:
                                        strMap.put("NOTE", value.replaceAll("\r|\n*",""));
                                        break;
                                    case 3:
                                        strMap.put("PORT", replacernNum(value,reg));
                                        break;
                                    case 4:
                                        strMap.put("PORT_TYPE", replacern(value));
                                        break;
                                    case 5:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("IN_PRESON_1", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("IN_DRIVERS", numMap.get("IN_PRESON_1"));
                                        break;
                                    case 6:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("IN_PRESON_2", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        break;
                                    case 7:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("OUT_PRESON_1", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        numMap.put("OUT_DRIVERS", numMap.get("OUT_PRESON_1"));
                                        break;
                                    case 8:
                                        value = value.replaceAll("[^0-9]+", "").replaceAll("[^\\d]+", "");
                                        numMap.put("OUT_PRESON_2", "".equals(value)?0:Double.parseDouble(value.replaceAll("[^\\d]+", "")));
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            case 0:
                                //得到单元格内数字
                                Double dvalue = c.getNumberValue();
                                switch (j) {
                                    case 1:
                                        strMap.put("AREA", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 2:
                                        strMap.put("NOTE", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 3:
                                        strMap.put("PORT", (dvalue+"").replaceAll("\r|\n*",""));
                                        break;
                                    case 4:
                                        strMap.put("PORT_TYPE", dvalue+"");
                                        break;
                                    case 5:
                                        numMap.put("IN_PRESON_1", dvalue);
                                        numMap.put("IN_DRIVERS", dvalue);
                                        break;
                                    case 6:
                                        numMap.put("IN_PRESON_2", dvalue);
                                        break;
                                    case 7:
                                        numMap.put("OUT_PRESON_1", dvalue);
                                        numMap.put("OUT_DRIVERS", dvalue);
                                        break;
                                    case 8:
                                        numMap.put("OUT_PRESON_2", dvalue);
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            default:
                        }
                    }
                }
                String area = strMap.get("AREA");
                tempSql.append((area == null || "".equals(area)) ? null + "," : "'" + area + "',");
                String port = strMap.get("PORT");
                if(port == null || "".equals(port)){
                    tempSql.append( null + "," );
                }else if(port.startsWith("蛇口")){
                    tempSql.append(  "'蛇口',");
                }else if(port.startsWith("惠州")){
                    tempSql.append(  "'惠州',");
                }else{
                    tempSql.append("'" + port + "',");
                }
                tempSql.append(strMap.get("PORT_TYPE") == null ? null + "," : "'" + strMap.get("PORT_TYPE") + "',");
                tempSql.append("'水运','1',");
                tempSql.append((numMap.get("IN_PRESON_1") == null ? 0 : numMap.get("IN_PRESON_1").intValue()) + (numMap.get("IN_PRESON_2") == null ? 0 : numMap.get("IN_PRESON_2").intValue()));
                tempSql.append(",");
                tempSql.append((numMap.get("OUT_PRESON_1") == null ? 0 : numMap.get("OUT_PRESON_1").intValue()) + (numMap.get("OUT_PRESON_2") == null ? 0 : numMap.get("OUT_PRESON_2").intValue()));
                tempSql.append(",");
                tempSql.append(numMap.get("IN_DRIVERS") == null ? 0 : numMap.get("IN_DRIVERS").intValue());
                tempSql.append(",");
                tempSql.append(numMap.get("OUT_DRIVERS") == null ? 0 : numMap.get("OUT_DRIVERS").intValue());
                tempSql.append(",");
                String note = strMap.get("NOTE");
                tempSql.append((note == null || "".equals(note)) ? null + "," : "'" + note + "',");
                tempSql.append("to_date('" + dirName + "','yyyyMMdd'),");
                tempSql.append("to_date('" + dirName + "','yyyyMMdd')");
                tempSql.append(SQL_SUF);
                datas.add(tempSql.toString());
                System.out.println(tempSql.toString());
            }
            suc = FileUtil.makeSql(datas, path, dirName, f.getName());
            //关闭输入流
            fi.close();
        } catch (Exception e){
            e.printStackTrace();
            System.out.println("解析Excel失败");
            suc = false;
        }finally {
            try{
                if(fi!=null){
                    fi.close();
                }
            }catch (Exception e){
                suc = false;
            }
            return  suc;
        }
    }

    /**
     * Method description :
     * 拆分合并单元格，可以从第几行第几列开始
     * sheet：sheet页
     * col：第几行第几列  例如A1
     * Author：  liuzhuang
     */
    private static void removeMerge(Sheet sheet, String col) {
        if (col == null || "".equals(col)) {
            col = "A1";
        }
        CellReference ref = new CellReference(col);
        //遍历sheet中的所有的合并区域
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            String value = "";
            CellRangeAddress region = sheet.getMergedRegion(i);
            Row firstRow = sheet.getRow(region.getFirstRow());
            Cell firstCellOfFirstRow = firstRow.getCell(region.getFirstColumn());
            //如果第一个单元格的是字符串
            if (firstCellOfFirstRow.getCellType() == Cell.CELL_TYPE_STRING) {
                value = firstCellOfFirstRow.getStringCellValue();
            }
            //判断到C1才进行拆分单元格
            if (region.getFirstRow() == ref.getRow() && region.getLastColumn() == ref.getCol()) {
                sheet.removeMergedRegion(i);
            }
            //设置第一行的值为，拆分后的每一行的值
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (region.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellValue(value);
                    }
                }
            }
        }
    }
    /**
     * 判断单元格是否为空
     */
    public static boolean isNull(Cell cell) {
        if(cell == null){
            return true;
        }else{
            String str = cell.getStringCellValue();
            return str == null || str.equals("") || str.equals("0") ||str.equals("0.0");
        }
    }
    public static boolean isNullStr(String str) {
        return str == null || str.equals("") || str.equals("0") ||str.equals("0.0");
    }
    /**
     * 过滤掉 \r\n
     */
    private static String replacern(String value){
        return value.replaceAll("\r|\n*","");
    }
    /**
     * 过滤掉 \r\n和数字和（） () 和其他字符
     */
    private static String replacernNum(String value,String reg){
        return value.replaceAll(reg, "");
    }
    /**
     * 判断单元格是否为空
     */
    private static String replacernd(String value){
        return value.replaceAll("\r|\n*","").replaceAll("[^\\d.]+","");
    }

    /**
     * double 转 int 去掉 .0
     * @param d
     * @return
     */
    private static String double2int(double d){
        String a = String.valueOf(d);
        if(a.endsWith(".0")){
            String b = a.substring(0,a.length()-2);
            return b;
        }
        return a;
    }
    /**
     * 获取 读取起始位置
     */
    public static int getStart(Sheet sheet) {
        for (int i = 0; i < 80; i++) {
            Cell cell = sheet.getRow(i).getCell(1);
            if (cell != null) {
                cell.setCellType(Cell.CELL_TYPE_STRING);
            }
            String cellValue = cell.getStringCellValue();
            if ( cellValue!=null &&cellValue.startsWith("辽宁")) {
                return i;
            }

        }
        return 5;
    }
}
