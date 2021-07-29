package com.github.donghyeok.excel;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.util.List;

class SimpleExcelWriter<T> {
    protected final Class<T> tClass;

    public SimpleExcelWriter(Class<T> tClass) {
        this.tClass = tClass;
    }

    public SXSSFWorkbook writeObjectsToWorkbook(final List<T> objects, final String sheetName) {
        return null;
    }

    public void writeObjectsToWorkbookAndDownload(final List<T> objects, String sheetName, final HttpServletResponse response) {
        SXSSFWorkbook workbook = writeObjectsToWorkbook(objects, sheetName);
    }

}
