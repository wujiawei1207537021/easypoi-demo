package com.yuntsoft.easypoidemo.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Word {

    private String title ;

    private String content;

    private String department;

    private String personName;

    private LocalDateTime time;

    private String imagesUrl;




}
