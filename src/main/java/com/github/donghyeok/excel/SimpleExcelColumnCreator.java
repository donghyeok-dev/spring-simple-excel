package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
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

    public Cell createBodyCell(Cell cell) {
        cell.setCellStyle(bodyCellStyle);
        return cell;
    }
}
