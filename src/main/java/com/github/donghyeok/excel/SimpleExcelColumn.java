package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.ExcelColumn;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;

@Getter
@Setter
public class SimpleExcelColumn {
    private Field field;
    private ExcelColumn excelColumn;

    @Builder
    public SimpleExcelColumn(Field field, ExcelColumn excelColumn) {
        this.field = field;
        this.excelColumn = excelColumn;
    }
}
