package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import com.github.donghyeok.excel.example.SampleDto;
import com.github.donghyeok.excel.exception.ExcelWriterException;
import com.github.drapostolos.typeparser.TypeParser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
    Map<Integer, SimpleExcelColumnCreator> columnOrderMap;
    SXSSFWorkbook workbook;
    int rowIdx;
    int colIdx;

    public SimpleExcelWriter(Class<T> tClass) {
        this.tClass = tClass;
        this.typeParser = TypeParser.newBuilder().build();
        this.workbook = new SXSSFWorkbook(100);
        this.workbook.setCompressTempFiles(true);
        this.comparator = Integer::compareTo; // 오름차순, 내림차순: comparator = (s1, s2)->s2.compareTo(s1);
        this.columnOrderMap = new TreeMap<>(comparator);
        this.rowIdx = 0;
        this.colIdx = 0;

        ReflectionUtils.doWithFields(SampleDto.class, f -> {
            final SimpleExcelColumn simpleExcelColumn = f.getAnnotation(SimpleExcelColumn.class);
            if(simpleExcelColumn != null) {

                ReflectionUtils.makeAccessible(f);
                this.columnOrderMap.put(simpleExcelColumn.columnOrder(), SimpleExcelColumnCreator.builder()
                        .workbook(this.workbook)
                        .field(f)
                        .simpleExcelColumn(simpleExcelColumn)
                        .build());
            }
        });
    }

    public void setHeader(SXSSFSheet sheet) {
        Row headerRow = sheet.createRow(this.rowIdx++);
        headerRow.setHeight((short)500);

        int idx = this.colIdx;
        for(Map.Entry<Integer, SimpleExcelColumnCreator> el : columnOrderMap.entrySet()) {
            SimpleExcelColumnCreator creator = el.getValue();
            sheet.setColumnWidth(idx, creator.getSimpleExcelColumn().columnWidth());
            creator.createHeaderCell(headerRow.createCell(idx));
            idx++;
        }
    }

    public void addRow(SXSSFSheet sheet, T object) {
        try {
            Row row = sheet.createRow(this.rowIdx++);
            int idx = this.colIdx;
            row.setHeight((short)500);

            for(Map.Entry<Integer, SimpleExcelColumnCreator> el : columnOrderMap.entrySet()) {
                SimpleExcelColumnCreator creator = el.getValue();
                Field field = creator.getField();
                Object value = field.get(object);
                creator.createBodyCell(row.createCell(idx), value);
                idx++;
            }
        } catch (Exception e) {
            throw new ExcelWriterException(e.getMessage(), e);
        }
    }

    public SXSSFWorkbook writeObjectsToWorkbook(final List<T> objects, final String sheetName) {
        SXSSFSheet sheet = workbook.createSheet(StringUtils.hasText(sheetName) ? sheetName : "sheet1");
        setHeader(sheet);
        objects.forEach(t -> addRow(sheet, t));

        // footer 넣기
        int idx = this.colIdx;
        for(Map.Entry<Integer, SimpleExcelColumnCreator> el : columnOrderMap.entrySet()) {
            SimpleExcelColumnCreator creator = el.getValue();
            System.out.println(creator.getSimpleExcelColumn().headerName() + ": " + creator.getSum());
            idx++;
        }
        
        return this.workbook;
    }

    public void writeObjectsToWorkbookAndDownload(final List<T> objects, String sheetName, final HttpServletResponse response) {
    }

}
