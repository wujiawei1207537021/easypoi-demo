package com.yuntsoft.easypoidemo.utils;

import cn.afterturn.easypoi.excel.annotation.Excel;
import com.yuntsoft.easypoidemo.entity.Person;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class AnnotationUtils {


    /***
     * 测试
     * @param args
     */
    public static void main(String[] args) throws InstantiationException, IllegalAccessException, NoSuchFieldException {
        List<String> stringList = new ArrayList<>();
        stringList.add("photo");
        Class clazz = changeAnnotation(Person.class,stringList);
        Class clazz2 = changeAnnotation(Person.class,stringList);

    }

    /***
     * 测试修改注解
     */
    public static <T> Class changeAnnotation(Class<T> clazz, List<String> hideList) {
        Field[] fields = clazz.getDeclaredFields();
        if (fields == null || fields.length < 1) {
            return clazz;
        }
        for (Field field : fields) {
            Excel excel = field.getAnnotation(Excel.class);
                InvocationHandler invocationHandler = Proxy.getInvocationHandler(excel);
                Field declaredField = null;
                try {
                    declaredField = invocationHandler.getClass().getDeclaredField("memberValues");
                } catch (NoSuchFieldException e) {
                    e.printStackTrace();
                }
                // 因为这个字段事 private final 修饰，所以要打开权限
                declaredField.setAccessible(true);
                // 获取 memberValues
                Map memberValues = null;
                try {
                    memberValues = (Map) declaredField.get(invocationHandler);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            if (hideList.indexOf(excel.name()) != -1) {
                // 修改 value 属性值
                memberValues.put("isColumnHidden", true);
                System.out.println(excel.isColumnHidden());
            }else {
                memberValues.put("isColumnHidden", false);
            }
        }
        return clazz;
    }
}
