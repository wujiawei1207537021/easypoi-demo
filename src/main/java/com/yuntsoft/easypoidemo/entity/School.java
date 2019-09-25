package com.yuntsoft.easypoidemo.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

@Data
public class School {

    @Excel(name = "学校名称")
    private String name ;
    @Excel(name = "学校学生人数")
    private String studentNum;
    @Excel(name = "学校教师人数")
    private String teacherNum;

}
