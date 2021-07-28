package com.github.donghyeok.excel.annotation;

import com.github.donghyeok.excel.enums.ExcelDataAlignment;
import com.github.donghyeok.excel.enums.ExcelDataType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
    String headerName() default "";
    ExcelDataType type() default ExcelDataType.STRING;
    ExcelDataAlignment alignment() default ExcelDataAlignment.CENTER;
    int width() default 4096;
}
