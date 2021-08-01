package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.ExcelColumn;
import com.github.donghyeok.excel.example.SampleDto;
import com.github.donghyeok.excel.exception.ExcelWriterException;
import com.github.drapostolos.typeparser.TypeParser;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.util.ReflectionUtils;
import org.springframework.util.StringUtils;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.*;

class SimpleExcelWriter<T> {
    protected final Class<T> tClass;
    protected final TypeParser typeParser;
    Comparator<Integer> comparator;
    Map<Integer, SimpleExcelColumn> headerOrderMap;
    SXSSFWorkbook workbook;
    int rowIdx;
    int colIdx;

    public SimpleExcelWriter(Class<T> tClass) {
        this.tClass = tClass;
        this.typeParser = TypeParser.newBuilder().build();
        this.workbook = new SXSSFWorkbook(100);
        this.workbook.setCompressTempFiles(true);
        this.comparator = Integer::compareTo; // 오름차순, 내림차순: comparator = (s1, s2)->s2.compareTo(s1);
        this.headerOrderMap = new TreeMap<>(comparator);
        this.rowIdx = 0;
        this.colIdx = 0;

        ReflectionUtils.doWithFields(SampleDto.class, f -> {
            final ExcelColumn excelColumn = f.getAnnotation(ExcelColumn.class);
            if(excelColumn != null) {
                ReflectionUtils.makeAccessible(f);
                this.headerOrderMap.put(excelColumn.headerOrder(), SimpleExcelColumn.builder()
                        .field(f)
                        .excelColumn(excelColumn)
                        .build());
            }
        });
    }

    public CellStyle setHeaderStyle() {
        Font font = this.workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short)12);
        font.setBold(true);

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    public void setHeader(SXSSFSheet sheet) {
        Row headerRow = sheet.createRow(this.rowIdx++);
        headerRow.setHeight((short)500);
        CellStyle headerCellStyle = setHeaderStyle();

        int startIdx = this.colIdx;
        for(Map.Entry<Integer, SimpleExcelColumn> el : headerOrderMap.entrySet()) {
            Cell cell = headerRow.createCell(startIdx++);
            cell.setCellValue(el.getValue().getExcelColumn().headerName());
            cell.setCellStyle(headerCellStyle);
        }
    }

    public CellStyle setBodyStyle(SimpleExcelColumn excelColumn) {
        Font font = this.workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short)10);
        font.setBold(false);

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setFont(font);
        cellStyle.setAlignment(excelColumn.getExcelColumn().alignment());
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    public void addRow(SXSSFSheet sheet, T object) {
        try {
            Row row = sheet.createRow(this.rowIdx++);
            int idx = this.colIdx;
            row.setHeight((short)500);

            for(Map.Entry<Integer, SimpleExcelColumn> el : headerOrderMap.entrySet()) {
                SimpleExcelColumn excelColumn = el.getValue();
                Field field = excelColumn.getField();
                Object value = field.get(object);
                sheet.setColumnWidth(idx, excelColumn.getExcelColumn().width());
                Cell cell = row.createCell(idx++);

                cell.setCellStyle(setBodyStyle(excelColumn));

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
        } catch (Exception e) {
            throw new ExcelWriterException(e.getMessage(), e);
        }
    }

    public SXSSFWorkbook writeObjectsToWorkbook(final List<T> objects, final String sheetName) {
        SXSSFSheet sheet = workbook.createSheet(StringUtils.hasText(sheetName) ? sheetName : "sheet1");
        setHeader(sheet);
        objects.forEach(t -> addRow(sheet, t));
        return this.workbook;
    }

    public void writeObjectsToWorkbookAndDownload(final List<T> objects, String sheetName, final HttpServletResponse response) {
    }

}
