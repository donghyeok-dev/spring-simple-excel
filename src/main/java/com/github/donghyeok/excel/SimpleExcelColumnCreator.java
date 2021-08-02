package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import com.github.donghyeok.excel.exception.ExcelWriterException;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
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

    @Builder
    public SimpleExcelColumnCreator(SXSSFWorkbook workbook, Field field, SimpleExcelColumn simpleExcelColumn) {
        this.field = field;
        this.simpleExcelColumn = simpleExcelColumn;
        this.headerCellStyle = workbook.createCellStyle();
        this.headerFont = workbook.createFont();
        this.bodyCellStyle = workbook.createCellStyle();
        this.bodyFont = workbook.createFont();

        if(simpleExcelColumn.includeFooterSum() && !isNumberic(this.field.getType())) {
            throw new ExcelWriterException(String.format("footer sum으로 지정된 %s 컬럼은 숫자 타입이 아닙니다.",
                    this.getSimpleExcelColumn().headerName()));
        }

        setHeaderStyle();
        setBodyStyle();
    }

    protected boolean isNumberic(Class<?> tClass) {
        return (Integer.class.equals(tClass)
                || int.class.equals(tClass)
                || Double.class.equals(tClass)
                || double.class.equals(tClass)
                || Float.class.equals(tClass)
                || float.class.equals(tClass));
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
        this.bodyCellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(this.simpleExcelColumn.bodyFormat()));
    }

    public void createHeaderCell(Cell cell) {
        cell.setCellValue(this.simpleExcelColumn.headerName());
        cell.setCellStyle(headerCellStyle);
    }

    public void createBodyCell(Cell cell, Object value) {
        cell.setCellStyle(bodyCellStyle);

        if(value == null) {
            cell.setCellValue((String) null);
        }else if(String.class.equals(field.getType())) {
            cell.setCellValue(value.toString());
        }else if(Integer.class.equals(field.getType()) || int.class.equals(field.getType())) {
            Integer curValue = (Integer) value;
            Integer sumValue = this.sum == null ? 0 : (Integer) this.sum;
            cell.setCellValue(curValue);
            this.sum = sumValue + curValue;
        }else if(Double.class.equals(field.getType()) || double.class.equals(field.getType())) {
            Double curValue = (Double) value;
            Double sumValue = this.sum == null ? 0.0 : (Double) this.sum;
            cell.setCellValue(curValue);
            this.sum = sumValue + curValue;
        }else if(Float.class.equals(field.getType()) || float.class.equals(field.getType())) {
            Float curValue = (Float) value;
            Float sumValue = this.sum == null ? 0f : (Float) this.sum;
            cell.setCellValue(curValue);
            this.sum = sumValue + curValue;
        } else if(Date.class.equals(field.getType())) {
            cell.setCellValue((Date) value);
        }else if(LocalDateTime.class.equals(field.getType()) || OffsetDateTime.class.equals(field.getType())) {
            cell.setCellValue((LocalDateTime) value);
        }
    }
}
