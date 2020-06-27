package com.example.mybatistest.excel.easyExcel;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;

//使用easyexcel写数据，需要先建立表格对应的实体类
public class ExcelEntity {
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private Integer age;
    @ExcelProperty("分数")
    private Integer score;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;

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

    public Integer getScore() {
        return score;
    }

    public void setScore(Integer score) {
        this.score = score;
    }

    public String getIgnore() {
        return ignore;
    }

    public void setIgnore(String ignore) {
        this.ignore = ignore;
    }

    @Override
    public String toString() {
        return "WriteExcelEntity{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", score=" + score +
                ", ignore='" + ignore + '\'' +
                '}';
    }
}
