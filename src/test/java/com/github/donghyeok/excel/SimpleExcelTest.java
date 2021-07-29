package com.github.donghyeok.excel;

import com.github.donghyeok.excel.example.SampleDto;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.IOException;
import java.util.List;
import static org.junit.jupiter.api.Assertions.*;

class SimpleExcelTest {

    @Test
    @DisplayName("test excel")
    void testexcel() throws IOException {
        // given
        SimpleExcel simpleExcel = new SimpleExcel();
        ClassPathResource resource = new ClassPathResource("excel/sample_upload_file.xlsx");

        // when
        List<SampleDto> results = simpleExcel.readExcelFromInputStream(SampleDto.class, resource.getInputStream());

        // then
        assertEquals(4, results.size());
        assertEquals(results.get(0).getNo(), 1121967);
        assertEquals(results.get(0).getCompanyName(), "Google");
        assertEquals(results.get(0).getCeoName(), "Gman");
        assertEquals(results.get(0).getEmail(), "Gman@gmail.com");
        assertEquals(results.get(0).getPhone(), "010-1111-7998");
        assertNull(results.get(0).getAddress());

        assertEquals(results.get(3).getNo(), 5047311);
        assertEquals(results.get(3).getCompanyName(), "AWS");
        assertEquals(results.get(3).getCeoName(), "Aman");
        assertEquals(results.get(3).getEmail(), "Aman@naver.com");
        assertEquals(results.get(3).getPhone(), "010-4444-6945");
    }
}