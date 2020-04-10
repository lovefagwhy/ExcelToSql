package com.excel;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

/**
 * @Author: LiuZhuang
 * Description:
 * Date:Created in 20:43 2020/4/8.
 * Modified By:
 */
public class PropertyUtil {
    //查询参数表
    public static Map<String, String> getPortProps(String path) throws Exception {
        System.out.println("                                                  ");
        Properties prop = new Properties();
        //读取属性文件a.properties
        System.out.println("解析配置文件port.properties");
        InputStream inputStream = new FileInputStream(path + "/port.properties");
        BufferedReader bf = new BufferedReader(new InputStreamReader(inputStream));
        //加载属性列表
        prop.load(inputStream);
        Iterator<String> it = prop.stringPropertyNames().iterator();
        Map<String, String> map = new HashMap<>();
        String key;
        String value;
        while (it.hasNext()) {
            key = it.next();
            value = prop.getProperty(key);
            key = new String(key.getBytes("iso-8859-1"), "GBK");
            map.put(key, new String(value.getBytes("iso-8859-1"), "GBK"));
            System.out.println(key + ":" + map.get(key));
        }
        System.out.println("                                                  ");
        inputStream.close();
        bf.close();
        return map;
    }

    //查询配置表
    public static Map<String, String> getConfigProps(String path) throws Exception {
        System.out.println("                                                  ");
        Properties prop = new Properties();
        //读取属性文件a.properties
        System.out.println("解析配置文件config.properties");
        InputStream inputStream = new FileInputStream(path + "/config.properties");
        BufferedReader bf = new BufferedReader(new InputStreamReader(inputStream));
        //加载属性列表
        prop.load(inputStream);
        Iterator<String> it = prop.stringPropertyNames().iterator();
        Map<String, String> map = new HashMap<>();
        String key;
        String value;
        while (it.hasNext()) {
            key = it.next();
            value = prop.getProperty(key);
            map.put(key, value);
            System.out.println(key + ":" + map.get(key));
        }
        System.out.println("                                                  ");
        inputStream.close();
        bf.close();
        return map;
    }
}
