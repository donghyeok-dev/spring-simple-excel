package com.github.donghyeok.excel;

import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class SimpleExcel {
    SimpleExcelFactory excelFactory;

    public SimpleExcel() {
        this.excelFactory = new SimpleExcelFactory();
    }

    public <T> List<T> readExcel(final Class<T> tClass, final MultipartFile multipartFile) throws IOException {
        return excelFactory.createExcelReader(tClass).convertInStreamToList(multipartFile.getInputStream());
    }

    public <T> List<T> readExcel(final Class<T> tClass, final FileInputStream fileInputStream) {
        return excelFactory.createExcelReader(tClass).convertInStreamToList(fileInputStream);
    }

    public <T> List<T> readExcel(final Class<T> tClass, final InputStream inputStream) {
        return excelFactory.createExcelReader(tClass).convertInStreamToList(inputStream);
    }


}
