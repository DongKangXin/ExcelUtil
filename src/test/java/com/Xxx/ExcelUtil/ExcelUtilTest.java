package com.Xxx.ExcelUtil;

import com.Xxx.ExcelUtil.entity.Student;
import org.apache.poi.ss.usermodel.DateUtil;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;


public class ExcelUtilTest {

    @Test
    public void parseExcel() {
        File file = new File("src/test/resources/Student.xlsx");
        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            Collection<Student> students = ExcelUtil.parseExcel(fileInputStream, Student.class);
            Assertions.assertNotNull(students);
            Assertions.assertEquals(students.size(),3);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        Date time = new GregorianCalendar(1900, 0, 0).getTime();
        Date dateTime = DateUtil.getJavaCalendar(0).getTime();
        Assertions.assertNotNull(time);
        Assertions.assertNotNull(dateTime);
    }

    public void testParseExcel() {
    }

    public void testParseExcel1() {
    }
}