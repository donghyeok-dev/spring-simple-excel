package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import com.github.donghyeok.excel.exception.ExcelReaderException;
import com.github.drapostolos.typeparser.TypeParser;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.BeanUtils;
import org.springframework.util.ReflectionUtils;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

class SimpleExcelReader<T> {
    protected final Class<T> tClass;
    protected final TypeParser typeParser;

    public SimpleExcelReader(Class<T> tClass) {
        this.tClass = tClass;
        this.typeParser = TypeParser.newBuilder().build();
    }

    public List<T> convertInStreamToList(final InputStream inputStream) {
        Map<String, Integer> headerNameMap = new HashMap<>();
        Map<Integer, Field> headerIndexMap = new HashMap<>();
        List<T> resultList = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            sheet.rowIterator().forEachRemaining(cells -> {
                if(cells.getRowNum() == 0) {
                    // header
                    cells.cellIterator().forEachRemaining(cell -> headerNameMap.put(cell.getStringCellValue(), cell.getColumnIndex()));

                    ReflectionUtils.doWithFields(tClass, f -> {
                        final SimpleExcelColumn simpleExcelColumn = f.getAnnotation(SimpleExcelColumn.class);
                        if(simpleExcelColumn != null && headerNameMap.containsKey(simpleExcelColumn.headerName()))
                            headerIndexMap.put(headerNameMap.get(simpleExcelColumn.headerName()), f);
                    });
                }else {
                    // body
                    final T object = BeanUtils.instantiateClass(tClass);

                    cells.cellIterator().forEachRemaining(cell -> {
                        Field f = headerIndexMap.get(cell.getColumnIndex());
                        setValue(object, cell, f);
                    });

                    resultList.add(object);
                }
            });

        } catch (Exception e) {
            throw new ExcelReaderException(e.getMessage(), e);
        }

        return resultList;
    }

    private String getValue(final Cell cell) {
        if(Objects.isNull(cell)) return null;

        String value;
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getLocalDateTimeCellValue().toString();
                }
                else {
                    value = BigDecimal.valueOf(cell.getNumericCellValue()).toPlainString();
                }
                if (value.endsWith(".0"))
                    value = value.substring(0, value.length() - 2);
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                value = String.valueOf(cell.getCellFormula());
                break;
            case ERROR:
                value = ErrorEval.getText(cell.getErrorCellValue());
                break;
            case BLANK:
            case _NONE:
            default:
                value = "";
        }
        return value;
    }

    private void setValue(Object object, final Cell cell, Field field) {
        try {
            field.setAccessible(true);
            field.set(object, typeParser.parseType(getValue(cell), field.getType()));
        } catch (IllegalAccessException e) {
            throw new ExcelReaderException(e.getMessage(), e);
        }
    }
}
