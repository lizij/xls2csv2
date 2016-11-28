package com.xlscsv.converter;

/**
 * Created by Lizij on 2016/11/28.
 */
public class TestMain {
    public static void main(String[] args) throws Exception{
        String basePath = System.getProperty("user.dir");
        Converter.convert(basePath + "\\input\\test1.csv", basePath + "\\output\\", -1);
        Converter.convert(basePath + "\\input\\test2.xls", basePath + "\\output\\", -1);
        Converter.convert(basePath + "\\input\\test3.xlsx", basePath + "\\output\\", -1);
    }
}
