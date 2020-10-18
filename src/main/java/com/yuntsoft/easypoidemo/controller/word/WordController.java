package com.yuntsoft.easypoidemo.controller.word;

import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.word.WordExportUtil;
import com.yuntsoft.easypoidemo.utils.WordlUtiles;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/word")
public class WordController {
    /**
     * 根据模板获取
     *
     * @param response
     */
    @GetMapping()
    public void getWord(HttpServletResponse response, @RequestParam(required = false, defaultValue = "测试标题") String title) throws Exception {

        Map<String, Object> map = new HashMap<String, Object>();
        map.put("title", title);
        map.put("department", "部门信息");
        map.put("personName", "人名称");
        map.put("time", new Date());
        map.put("content", "测试内容");
        ImageEntity image = new ImageEntity();
        image.setHeight(200);
        image.setWidth(500);
        image.setUrl("images.png");
        image.setType(ImageEntity.URL);
        map.put("image", image);
        XWPFDocument doc = WordExportUtil.exportWord07(
                "Image.docx", map);
        WordlUtiles.downLoadWord("文本.docx", response, doc);

    }

}
