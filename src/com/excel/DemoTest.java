package com.excel;

/**
 * Description:
 * Author:        liuzhuang
 * Create Date:   2020/4/15 14:16
 */
public class DemoTest {
    public static void main(String[] args){
        String string = "辽\n33,23宁/(12)!d..。，大@#$%^&*~\r\n天津（1）";
        String s = string.replaceAll("(\r|\n|[0-9]|[a-zA-Z]|（|）|\\.|\\,|\\，|\\。|/|\\(|\\)|#|\\^|@|\\$|%|&|\\*|!|\\~)*", "");
        System.out.println(s);
    }
}
