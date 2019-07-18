package com.Xxx.ExcelUtil.test;


import com.Xxx.ExcelUtil.ExcelUtil;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;

public class TestMain {


    public static void main(String[] args) {
        File file = new File("src/main/resources/Student.xlsx");
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            Collection<Student> students = ExcelUtil.parseExcel(fileInputStream, Student.class);
            System.out.println(students);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Date time = new GregorianCalendar(1900, 0, 0).getTime();
        System.out.println(DateUtil.getJavaCalendar(0).getTime());
    }
}
