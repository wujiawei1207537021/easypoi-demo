package com.yuntsoft.easypoidemo.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;

public class WordlUtiles {

    /**
     * word 导出
     *
     * @param fileName
     * @param response
     * @param document
     */
    public static void downLoadWord(String fileName, HttpServletResponse response, XWPFDocument document) {
        try {
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/msword");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            document.write(response.getOutputStream());
        } catch (IOException e) {
            //throw new NormalException(e.getMessage());
        }
    }


}