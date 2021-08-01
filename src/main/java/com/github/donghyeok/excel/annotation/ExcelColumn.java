package com.github.donghyeok.excel.annotation;

import com.github.donghyeok.excel.enums.ExcelDataType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
    String headerName() default "";
    int headerOrder() default 1;
    ExcelDataType type() default ExcelDataType.STRING;
    HorizontalAlignment alignment() default HorizontalAlignment.CENTER;
    int width() default 4096;
}
