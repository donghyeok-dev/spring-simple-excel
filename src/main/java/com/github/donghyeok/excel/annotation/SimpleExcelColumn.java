package com.github.donghyeok.excel.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface SimpleExcelColumn {
    /**
     * 헤더 컬럼명
     */
    String headerName() default "";

    /**
     * 엑셀파일을 생성할 때 컬럼 나열 순서
     */
    int columnOrder() default 1;

    /**
     * 컬럼 넓이
     */
    int columnWidth() default 4096;

    /**
     * Header Style
     */
    String headerFontName() default "맑은 고딕";
    short headerFontSize() default 12;
    boolean headerFontBold() default true;
    IndexedColors headerBackgroundColor() default IndexedColors.GREY_25_PERCENT;
    HorizontalAlignment headerAlignment() default HorizontalAlignment.CENTER;

    /**
     * Body Style
     */
    String bodyFontName() default "맑은 고딕";
    short bodyFontSize() default 10;
    boolean bodyFontBold() default false;
    IndexedColors bodyBackgroundColor() default IndexedColors.WHITE;
    HorizontalAlignment bodyAlignment() default HorizontalAlignment.CENTER;

    /**
     * 숫자 천단위 표시하려면 #,##0
     */
    String bodyFormat() default "";

    boolean includeFooterSum() default false;

}
