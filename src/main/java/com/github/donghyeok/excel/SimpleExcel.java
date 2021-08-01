package com.github.donghyeok.excel;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * spring Simple Excel
 *
 * This is a library that makes it easier to download processed data to Excel or to read Excel files to convert them into data objects.
 * The library development environment is spring boot 2.4.
 *
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 */
public class SimpleExcel {
    private final SimpleExcelFactory excelFactory;

    public SimpleExcel() {
        this.excelFactory = new SimpleExcelFactory();
    }

    /**
     MultipartFile로 전송된 엑셀 파일을 읽어서 지정된 Type으로 변환하여 반환한다.
     */
    public <T> List<T> readExcelFromMultipartFile(final Class<T> tClass, final MultipartFile multipartFile) throws IOException {
        return excelFactory.createExcelReader(tClass).convertInStreamToList(multipartFile.getInputStream());
    }

    /**
     InputStream의 엑셀 파일을 읽어서 지정된 Type으로 변환하여 반환한다.
     */
    public <T> List<T> readExcelFromInputStream(final Class<T> tClass, final InputStream inputStream) {
        return excelFactory.createExcelReader(tClass).convertInStreamToList(inputStream);
    }

    /**
     지정된 Type의 객체를 엑셀 파일로 변환하여 반환한다.
     */
    public <T> SXSSFWorkbook writeObjectsToWorkbook(final Class<T> tClass, final List<T> objects, final String sheetName) {
        return excelFactory.createExcelWriter(tClass).writeObjectsToWorkbook(objects, sheetName);
    }

    /**
     지정된 Type의 객체를 엑셀 파일로 변환하고 response로 파일을 내려준다.
     */
    public <T> void writeObjectsToWorkbookAndDownload(final Class<T> tClass, final List<T> objects, final String sheetName, HttpServletResponse response) {
        excelFactory.createExcelWriter(tClass).writeObjectsToWorkbookAndDownload(objects, sheetName, response);
    }
}
