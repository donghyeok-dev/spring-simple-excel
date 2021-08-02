package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import com.github.donghyeok.excel.exception.ExcelWriterException;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Date;

@Getter
@Setter
public class SimpleExcelColumnCreator {
    private Field field;
    private SimpleExcelColumn simpleExcelColumn;
    private CellStyle headerCellStyle;
    private Font headerFont;
    private CellStyle bodyCellStyle;
    private Font bodyFont;
    private Object sum;
    private Calculation addition;

    @Builder
    public SimpleExcelColumnCreator(SXSSFWorkbook workbook, Field field, SimpleExcelColumn simpleExcelColumn) {
        this.field = field;
        this.simpleExcelColumn = simpleExcelColumn;
        this.headerCellStyle = workbook.createCellStyle();
        this.headerFont = workbook.createFont();
        this.bodyCellStyle = workbook.createCellStyle();
        this.bodyFont = workbook.createFont();

        setHeaderStyle();
        setBodyStyle();

        if(Integer.class.equals(this.field.getType()) || int.class.equals(this.field.getType()) ) {
            this.sum = 0;
            addition = (Calculation<Integer>) Integer::sum;
        }else if(Double.class.equals(this.field.getType()) || double.class.equals(this.field.getType()) ) {
            this.sum = 0.0;
            addition = (Calculation<Double>) Double::sum;
        }else if(Float.class.equals(this.field.getType()) || float.class.equals(this.field.getType()) ) {
            this.sum = 0f;
            addition = (Calculation<Float>) Float::sum;
        }
    }

    protected void setHeaderStyle() {
        this.headerFont.setFontName(this.simpleExcelColumn.headerFontName());
        this.headerFont.setFontHeightInPoints(this.simpleExcelColumn.headerFontSize());
        this.headerFont.setBold(this.simpleExcelColumn.headerFontBold());

        this.headerCellStyle.setBorderBottom(BorderStyle.THIN);
        this.headerCellStyle.setBorderLeft(BorderStyle.THIN);
        this.headerCellStyle.setBorderRight(BorderStyle.THIN);
        this.headerCellStyle.setBorderTop(BorderStyle.THIN);
        this.headerCellStyle.setFont(this.headerFont);
        this.headerCellStyle.setAlignment(this.simpleExcelColumn.headerAlignment());
        this.headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        this.headerCellStyle.setFillForegroundColor(this.simpleExcelColumn.headerBackgroundColor().getIndex());
        this.headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    protected void setBodyStyle() {
        this.bodyFont.setFontName(this.simpleExcelColumn.bodyFontName());
        this.bodyFont.setFontHeightInPoints(this.simpleExcelColumn.bodyFontSize());
        this.bodyFont.setBold(this.simpleExcelColumn.bodyFontBold());

        this.bodyCellStyle.setBorderBottom(BorderStyle.THIN);
        this.bodyCellStyle.setBorderLeft(BorderStyle.THIN);
        this.bodyCellStyle.setBorderRight(BorderStyle.THIN);
        this.bodyCellStyle.setBorderTop(BorderStyle.THIN);
        this.bodyCellStyle.setFont(this.bodyFont);
        this.bodyCellStyle.setAlignment(this.simpleExcelColumn.bodyAlignment());
        this.bodyCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        this.bodyCellStyle.setFillForegroundColor(this.simpleExcelColumn.bodyBackgroundColor().getIndex());
        this.bodyCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public void createHeaderCell(Cell cell) {
        cell.setCellValue(this.simpleExcelColumn.headerName());
        cell.setCellStyle(headerCellStyle);
    }

    protected boolean isNumberic(Class<?> tClass) {
        return (Integer.class.equals(tClass)
                || int.class.equals(tClass)
                || Double.class.equals(tClass)
                || double.class.equals(tClass)
                || Float.class.equals(tClass)
                || float.class.equals(tClass));
    }

    public void createBodyCell(Cell cell, Object value) {
        cell.setCellStyle(bodyCellStyle);

        if(simpleExcelColumn.includeFooterSum()) {
            if(!isNumberic(this.field.getType())) {
                System.out.println(this.getSimpleExcelColumn().headerName() + " " + this.field.getType());
                throw new ExcelWriterException("foot sum 컬럼은 숫자형타입만 가능합니다.");
            }
            this.sum = this.addition.apply(value, this.sum);
        }

        if(value == null) {
            cell.setCellValue((String) null);
        }else if(String.class.equals(field.getType())) {
            cell.setCellValue(value.toString());
        }else if(Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            cell.setCellValue((Integer) value);
        }else if(Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            cell.setCellValue((Double) value);
        }else if(Float.class.equals(field.getType()) || float.class.equals(field.getType())) {
            cell.setCellValue((Float) value);
        } else if(Date.class.equals(field.getType())) {
            cell.setCellValue((Date) value);
        }else if(LocalDateTime.class.equals(field.getType()) || OffsetDateTime.class.equals(field.getType())) {
            cell.setCellValue((LocalDateTime) value);
        }
    }
}

@FunctionalInterface
interface Calculation<T> {
    T apply(T x, T y);
}
