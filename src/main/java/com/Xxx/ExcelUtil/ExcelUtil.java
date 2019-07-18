package com.Xxx.ExcelUtil;


import com.Xxx.ExcelUtil.annotation.Header;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelUtil {

    private static final Date TIME_SINCE = new GregorianCalendar(1900, 0, 0).getTime();

    /*
     * 导出类
     */
    private static Class<?> clazz;

    /**
     * 导出类的所有属性
     */
    private static Field[] fields;

    /**
     * 属性值索引
     */
    private static Integer[] index;



    public static <T> Collection<T> parseExcel(InputStream inputStream, Class<T> clazz) {
        Map<String, String> fieldValueMap = getFieldValueMap(clazz);
        return parseExcel(inputStream, fieldValueMap, clazz);
    }

    public static <T> Collection<T> parseExcel(InputStream inputStream, String fieldValue, Class<T> clazz) {
        Map<String, String> fieldValueMap = getFieldValueMap(fieldValue);
        return parseExcel(inputStream, fieldValueMap, clazz);
    }

    public static <T> Collection<T> parseExcel(InputStream inputStream, Map<String, String> fieldValueMap, Class<T> clazz) {
        Collection<T> objects = new ArrayList<T>();
        setClazz(clazz);
        List<Sheet> sheets = readInputStream(inputStream);
        Sheet rows = sheets.get(0);
        Row row = rows.getRow(0);
        Map<String, Integer> valueIndexMap = getValueIndexMap(row);
        setIndex(fieldValueMap, valueIndexMap);
        int i = 1;
        row = rows.getRow(i);
        AccessibleObject.setAccessible(fields,true);
        while (row != null) {
            T object = getObject(row,clazz);
            objects.add(object);
            i++;
            row = rows.getRow(i);
        }
        AccessibleObject.setAccessible(fields,false);
        return objects;
    }

    /**
     * 读取输入流获取sheet对象集合
     *
     * @param inputStream
     * @return
     */
    private static List<Sheet> readInputStream(InputStream inputStream) {

        Workbook workbook = null;
        List<Sheet> sheets = new ArrayList<Sheet>();
        try {
            workbook = WorkbookFactory.create(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sheets.add(workbook.getSheetAt(i));
        }
        return sheets;
    }

    /**
     * 通过属性名与表头名的对应关系与表头名与表头索引的对应关系，建立属性值索引
     *
     * @param fieldValueMap
     * @param valueIndexMap
     */
    private static void setIndex(Map<String, String> fieldValueMap, Map<String, Integer> valueIndexMap) {
        index = new Integer[fields.length];
        for (int i = 0; i < fields.length; i++) {
            index[i] = valueIndexMap.get(fieldValueMap.get(fields[i].getName()));
        }
    }

    /**
     * 生成表头名与列数的对应关系
     *
     * @param row
     * @return
     */
    private static Map<String, Integer> getValueIndexMap(Row row) {
        Map<String, Integer> map = new HashMap<String, Integer>();
        Iterator<Cell> cellIterator = row.cellIterator();
        int i = 0;
        while (cellIterator.hasNext()) {
            map.put(cellIterator.next().toString(), i);
            i++;
        }
        return map;
    }

    /**
     * 通过key-value生成属性名与表头的对应关系
     *
     * @param keyValue
     * @return
     */
    private static Map<String, String> getFieldValueMap(String keyValue) {
        String[] keyValues = keyValue.split(",");
        Map<String, String> map = new HashMap<String, String>();
        for (int i = 0; i < keyValues.length; i++) {
            String[] split = keyValues[i].split(":");
            if (split.length == 2) {
                map.put(split[0], split[1]);
            } else {
                new Exception().printStackTrace();
            }
        }
        return map;
    }

    /**
     * 通过key-value生成属性名与表头的对应关系
     *
     * @param keyValue
     * @return
     */
    private static Map<String, String> getFieldValueMap(Map<String, String> keyValue) {
        return keyValue;
    }

    /**
     * 通过注解生成属性名与表头关系
     * @param clazz
     * @return
     */
    private static Map<String, String> getFieldValueMap(Class<?> clazz){
        Map<String,String> map = new HashMap<String, String>();
        Field[] fields = clazz.getDeclaredFields();
        for(int i = 0; i<fields.length;i++){
            String value = fields[i].getAnnotation(Header.class).value();
            map.put(fields[i].getName(),value);
        }
        return map;
    }

    /**
     * 初始化类Class与类属性Feild
     *
     * @param clazz1
     */
    public static void setClazz(Class<?> clazz1) {
        clazz = clazz1;
        fields = clazz.getDeclaredFields();
    }

    /**
     * 创建对象并赋值
     *
     * @param row 从表中获取的单行值
     * @return 返回创建的对象
     */
    private static <T> T getObject(Row row,Class<T> tClass) {

        try {
            T object = tClass.newInstance();
            for (int i = 0; i < fields.length; i++) {
                if(index[i]!=null){
                    String cell = row.getCell(index[i]).toString();
                    Object valueElement = getValueElement(cell, fields[i].getType());
                    fields[i].set(object,valueElement);
                }
            }
            return object;
        } catch (InstantiationException e) {
            e.printStackTrace();
            return null;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            return null;
        }

    }

    /**
     * 根据传入的字符串与要转换的格式进行转换
     * @param s
     * @param tClass
     * @param <T>
     * @return
     */
    public static <T> T getValueElement(String s ,Class<T> tClass){
        if(tClass.equals(String.class)){
            return (T)s;
        }
        if(tClass.equals(Double.class)){
            return tClass.cast((Number)Double.parseDouble(s));
        }
        if(tClass.equals(Integer.class)){
            return tClass.cast(Integer.parseInt(s.split("\\.")[0]));
        }
        if(tClass.equals(Date.class)){
            return tClass.cast(DateUtil.getJavaDate(Double.parseDouble(s)));
        }
        return null;
    }
}
