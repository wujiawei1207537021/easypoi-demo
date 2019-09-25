package com.yuntsoft.easypoidemo.controller.excl;

import com.yuntsoft.easypoidemo.entity.School;
import com.yuntsoft.easypoidemo.utils.ExcelUtiles;
import com.yuntsoft.easypoidemo.utils.FastExcel;
import io.swagger.annotations.Api;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.List;

@RestController
@Api("School")
@RequestMapping("/school")
public class CustomController {

    /***
     *  导出模板
     * @param httpServletResponse
     */
    @GetMapping("/down")
    public void down(HttpServletResponse httpServletResponse) throws IOException {
        InputStream is =  Thread.currentThread().getContextClassLoader().getResourceAsStream("school.xls");
        Workbook workbook = WorkbookFactory.create(is);
        ExcelUtiles.downLoadExcel1("学校.xls", httpServletResponse, workbook);
    }

    /**
     * 自定义一工具导入
     *
     * @param multipartFile
     * @return
     */
    @PostMapping("/imp1")
    public String import1(@RequestBody MultipartFile multipartFile) {
        List<School> schools = FastExcel.praseExcel(multipartFile, School.class);
        return schools.toString();
    }


}
