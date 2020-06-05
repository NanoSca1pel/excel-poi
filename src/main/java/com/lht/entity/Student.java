package com.lht.entity;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

/**
 * @author lhtao
 * @date 2020/6/5 10:05
 */
@Data
public class Student {

    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("生日")
    private Date birth;

    @ExcelProperty("年龄")
    private Integer age;

    /** 忽略这个字段 */
    @ExcelIgnore
    private String gender;
}
