package com.yuntsoft.easypoidemo.entity;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Person {
    public Person(String name, Integer sex, Date birthday, Date registrationDate) {
        this.name = name;
        this.sex = sex;
        this.birthday = birthday;
        this.registrationDate = registrationDate;
    }

    @Excel(name = "姓名")
    private String name;

    @Excel(name = "学生性别", replace = {"男_1", "女_0"}, suffix = "生", isImportField = "true_st")
    private Integer sex;

    @Excel(name = "出生日期", databaseFormat = "yyyyMMddHHmmss", format = "yyyy-MM-dd", isImportField = "true_st", width = 20)
    private Date birthday;

    @Excel(name = "进校日期", databaseFormat = "yyyyMMddHHmmss", format = "yyyy-MM-dd")
    private Date registrationDate;

    @Excel(name = "photo" ,type = 2 ,width = 40 , height = 20,imageType = 1)
    private String  photo;

}