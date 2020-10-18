package com.yuntsoft.easypoidemo.utils;

import cn.afterturn.easypoi.excel.annotation.Excel;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

@Deprecated
public class FastExcel {

    /**
     * 根据路径解析
     *
     * @param path
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> praseExcel(String path, Class<T> clazz) {
        File file = new File(path);
//       return this.praseExcel(file,clazz);
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return praseExcel(workbook, clazz);
    }

    /**
     * MultipartFile
     *
     * @param file
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> praseExcel(MultipartFile file, Class<T> clazz) {
        InputStream is = null;
        try {
            is = file.getInputStream();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return praseExcel(is, clazz);
    }

    /**
     * 根据输入输出流解析
     *
     * @param is
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> praseExcel(InputStream is, Class<T> clazz) {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return praseExcel(workbook, clazz);
    }

    public static <T> List<T> praseExcel(Workbook workbook, Class<T> clazz) {
        Sheet sheet = workbook.getSheetAt(0);
        List<T> rst = new ArrayList<>();
        if (sheet == null) {
            return rst;
        }
        int firstRowNum = sheet.getFirstRowNum();
        Row row = sheet.getRow(firstRowNum);
        short lastCellNum = row.getLastCellNum();
        // key:表头,value:对应的列数
        Map<String, Integer> cellNames = getCellMapping(row, lastCellNum);
        // key:映射的表头名字,value:对应的字段
        Map<String, Field> annotations = getFeildMapping(clazz);
        int lastRowNum = sheet.getLastRowNum();
        Set<String> keys = cellNames.keySet();
        try {
            for (int rowIndex = (++firstRowNum); rowIndex <= lastRowNum; rowIndex++) {
                T inst = clazz.newInstance();
                Row r = sheet.getRow(rowIndex);
                for (String key : keys) {
                    Field field = annotations.get(key);
                    if (field == null) {
                        continue;
                    }
                    Integer col = cellNames.get(key);
                    Cell cel = r.getCell(col);
                    if (cel == null) {
                        continue;
                    }
                    field.setAccessible(true);
                    String val = cel.getStringCellValue();
                    field.set(inst, val);
                }
                rst.add(inst);
            }
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return rst;
    }

    /**
     * 获取表头和列的映射关系
     *
     * @param row
     * @param lastCellNum
     * @return
     */
    private static Map<String, Integer> getCellMapping(Row row, short lastCellNum) {
        // key:表头,value:对应的列数
        Map<String, Integer> cellNames = new HashMap<>();
        Cell cell;
        for (int col = 0; col < lastCellNum; col++) {
            cell = row.getCell(col);
            String val = cell.getStringCellValue();
            cellNames.put(val, col);
        }
        return cellNames;
    }

    /**
     * 获取对象字段和Excel表头的字段映射关联
     *
     * @param clazz
     * @return
     */
    private static <T> Map<String, Field> getFeildMapping(Class<T> clazz) {
        // key:映射的表头名字,value:对应的字段
        Map<String, Field> annotations = new HashMap<>();
        Field[] fields = clazz.getDeclaredFields();
        if (fields == null || fields.length < 1) {
            return annotations;
        }
        for (Field field : fields) {
            Excel mapping = field.getAnnotation(Excel.class);
            if (mapping == null) {
                annotations.put(field.getName(), field);
            } else {
                annotations.put(mapping.name(), field);
            }
        }
        return annotations;
    }

//    /***
//     * 测试
//     * @param args
//     */
//    public static void main(String[] args) throws InstantiationException, IllegalAccessException, NoSuchFieldException {
//
//       Class clazz= changeAnnotation(Person.class, new ArrayList<>());
//        Class clazz2= changeAnnotation(clazz, new ArrayList<>());
//
//    }
//
//    /***
//     * 测试修改注解
//     */
//    public static <T> Class   changeAnnotation(Class<T> clazz,List<String> hideList)  {
//        Field[] fields = clazz.getDeclaredFields();
//        if (fields == null || fields.length < 1) {
//            return clazz;
//        }
//        for (Field field : fields) {
//            Excel excel = field.getAnnotation(Excel.class);
//            if(hideList.indexOf(excel.name())!=-1){
//                InvocationHandler invocationHandler = Proxy.getInvocationHandler(excel);
//                Field declaredField = null;
//                try {
//                    declaredField = invocationHandler.getClass().getDeclaredField("memberValues");
//                } catch (NoSuchFieldException e) {
//                    e.printStackTrace();
//                }
//                // 因为这个字段事 private final 修饰，所以要打开权限
//                declaredField.setAccessible(true);
//                // 获取 memberValues
//                Map memberValues = null;
//                try {
//                    memberValues = (Map) declaredField.get(invocationHandler);
//                } catch (IllegalAccessException e) {
//                    e.printStackTrace();
//                }
//                // 修改 value 属性值
//                memberValues.put("isColumnHidden", true);
//                System.out.println(excel.isColumnHidden());
//            }
//
//        }
//        return clazz;
//    }


}