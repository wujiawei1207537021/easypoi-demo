package com.yuntsoft.easypoidemo.controller.excl;

import com.yuntsoft.easypoidemo.entity.Person;
import com.yuntsoft.easypoidemo.utils.AnnotationUtils;
import com.yuntsoft.easypoidemo.utils.ExcelUtiles;
import io.swagger.annotations.Api;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@RestController
@Api("Person")
@RequestMapping("/excl")
public class EasyPOIController {


    @GetMapping("/down")
    public void down(HttpServletResponse httpServletResponse) {
        String fileName = "导出模板.xls";
        Class clazz = AnnotationUtils.changeAnnotation(Person.class, new ArrayList<>());
        ExcelUtiles.exportExcel(new ArrayList<>(), "标题", "sheetName", clazz, LocalDateTime.now() + fileName, httpServletResponse);
    }

    /***
     * excl根据实体类 导出模板
     * @param httpServletResponse
     */
    @GetMapping("/down-noimags")
    public void downNoImages(HttpServletResponse httpServletResponse) {
        this.down(httpServletResponse, false);
    }

    /**
     * excl 包含images
     *
     * @param httpServletResponse
     * @throws FileNotFoundException
     */
    @GetMapping("/down-imags")
    public void downImages(HttpServletResponse httpServletResponse, @RequestParam(required = false, defaultValue = "false") Boolean hasImag) {
        this.down(httpServletResponse, true);
    }

    /**
     * 导出
     *
     * @param httpServletResponse
     * @throws FileNotFoundException
     */
    public void down(HttpServletResponse httpServletResponse, Boolean hasImag) {
        Class clazz = Person.class;
        String fileName = "";
        List<Person> people = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Person person = new Person("姓名" + i, i % 2, new Date(), new Date());
            if (hasImag) {
                person.setPhoto("images.png");
                fileName = "有图片导出测试文件.xls";
                clazz = AnnotationUtils.changeAnnotation(clazz, new ArrayList<>());
            } else {
                List<String> list = new ArrayList<>();
                list.add("photo");
                clazz = AnnotationUtils.changeAnnotation(clazz, list);
                fileName = "没有图片导出测试文件.xls";
            }
            people.add(person);
        }
        ExcelUtiles.exportExcel(people, "标题", "sheetName", clazz, LocalDateTime.now() + fileName, httpServletResponse);
    }

    /**
     * easypoi 导入
     *
     * @param multipartFile
     * @return
     */
    @PostMapping("/imp")
    public String import1(@RequestBody MultipartFile multipartFile) {
        List<Person> people = ExcelUtiles.importExcel(multipartFile, 1, 1, Person.class);
        System.out.println(people.size());
        return people.toString();
    }


}
