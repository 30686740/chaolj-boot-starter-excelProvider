package com.chaolj.core.excelProvider.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    // 列名
    public String name() default "";

    // 排序
    public int order() default 1;

    // 时间格式
    public String dateFormat() default "yyyy-MM-dd HH:mm:ss";

    // 数字格式
    public String numFormat() default "";

    // 后缀
    public String suffix() default "";

    // 列宽
    public int width() default 10;

    /**
     * 是否舍入模式：RoundingMode <br/>
     * 0:ROUND_UP, 1:ROUND_DOWN, 2:ROUND_CEILING, 3:ROUND_FLOOR,
     * 4:ROUND_HALF_UP(四舍五入), 5:ROUND_HALF_DOWN, 6:ROUND_HALF_EVEN, 7:ROUND_UNNECESSARY
     */
    int type() default 1;

    /**
     * 支持除法运算
     */
    String divide() default "";

    /**
     * 兼容除法运算小数位溢出（保留位数）
     */
    int place() default 2;
}
