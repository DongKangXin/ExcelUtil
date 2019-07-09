package com.Xxx.ExcelUtil;


import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.AccessibleObject;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelUtil {

    /**
     * 导出类
     */
    private static Class<?> clazz;

    /**
     * 导出类的所有属性
     */
    private static Field[] fields;

    /**
     * 属性类型
     */
    private static Integer[] type;

    /**
     * 属性值索引
     */
    private static Integer[] index;

    /**
     * 导出对象集合
     */
    private static Collection<Object> objects;


    public static Collection<Object> parseExcel(InputStream inputStream, String fieldValue, Class<?> clazz) {
        Map<String, String> fieldValueMap = getFieldValueMap(fieldValue);
        return parseExcel(inputStream, fieldValueMap, clazz);
    }

    public static Collection<Object> parseExcel(InputStream inputStream, Map<String, String> fieldValueMap, Class<?> clazz) {
        objects = new ArrayList<Object>();
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
            Object object = getObject(row);
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
        type = new Integer[fields.length];
        for (int i = 0; i < fields.length; i++) {
            index[i] = valueIndexMap.get(fieldValueMap.get(fields[i].getName()));
            Class<?> name = fields[i].getType();
            if (name.equals(byte.class) || name.equals(Byte.class)) {
                ExcelUtil.type[i] = 0;
            } else if (name.equals(short.class) || name.equals(Short.class)) {
                ExcelUtil.type[i] = 1;
            } else if (name.equals(int.class) || name.equals(Integer.class)) {
                ExcelUtil.type[i] = 2;
            } else if (name.equals(long.class) || name.equals(Long.class)) {
                ExcelUtil.type[i] = 3;
            } else if (name.equals(float.class) || name.equals(Float.class)) {
                ExcelUtil.type[i] = 4;
            } else if (name.equals(double.class) || name.equals(Double.class)) {
                ExcelUtil.type[i] = 5;
            } else if (name.equals(boolean.class) || name.equals(Boolean.class)) {
                ExcelUtil.type[i] = 6;
            } else if (name.equals(char.class) || name.equals(Character.class)) {
                ExcelUtil.type[i] = 7;
            } else if (name.equals(String.class)) {
                ExcelUtil.type[i] = 8;
            } else if (name.equals(Date.class)) {
                ExcelUtil.type[i] = 9;
            }

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
    private static Object getObject(Row row) {

        try {
            Object o = clazz.newInstance();
            for (int i = 0; i < fields.length; i++) {
                String cell = row.getCell(i).toString();
                switch (type[i]) {
                    case 0:
                        fields[i].set(o, Byte.parseByte(cell));
                        break;
                    case 1:
                        fields[i].set(o, Short.parseShort(cell));
                        break;
                    case 2:
                        fields[i].set(o, Integer.parseInt(cell));
                        break;
                    case 3:
                        fields[i].set(o, Long.parseLong(cell));
                        break;
                    case 4:
                        fields[i].set(o, Float.parseFloat(cell));
                        break;
                    case 5:
                        fields[i].set(o, Double.parseDouble(cell));
                        break;
                    case 6:
                        fields[i].set(o, Boolean.parseBoolean(cell));
                        break;
                    case 7:
                        fields[i].set(o, cell.charAt(0));
                        break;
                    case 8:
                        fields[i].set(o, cell);
                        break;
                    case 9:
                        fields[i].set(o, Date.parse(cell));
                        break;
                    default:
                        break;
                }
            }
            return o;
        } catch (InstantiationException e) {
            e.printStackTrace();
            return null;
        } catch (IllegalAccessException e) {
            e.printStackTrace();
            return null;
        }

    }
}
