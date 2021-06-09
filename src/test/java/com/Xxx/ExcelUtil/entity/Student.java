package com.Xxx.ExcelUtil.entity;

import com.Xxx.ExcelUtil.annotation.Header;

import java.util.Date;

public class Student {

    @Header("姓名")
    private String name;

    @Header("年龄")
    private Integer age;

    @Header("分数")
    private Double score;

    @Header("出生日期")
    private Date date;

    public Double getScore() {
        return score;
    }

    public void setScore(Double score) {
        this.score = score;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
