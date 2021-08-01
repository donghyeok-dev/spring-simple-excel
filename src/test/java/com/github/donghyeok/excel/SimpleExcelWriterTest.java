package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.ExcelColumn;
import com.github.donghyeok.excel.example.SampleDto;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.util.ReflectionUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

class SimpleExcelWriterTest {
    List<SampleDto> list = new ArrayList<>();

    @BeforeEach
    void setUp() {
        list.add(SampleDto.builder().no(1).ceoName("대표A").companyName("회사A").email("a@gamil.com").build());
        list.add(SampleDto.builder().no(2).ceoName("대표B").companyName("회사B").email("b@gamil.com").build());
        list.add(SampleDto.builder().no(3).ceoName("대표C").companyName("회사C").email("c@gamil.com").build());
        list.add(SampleDto.builder().no(4).ceoName("대표D").companyName("회사D").email("d@gamil.com").build());
        list.add(SampleDto.builder().no(5).ceoName("대표E").companyName("회사E").email("e@gamil.com").build());
    }

    @Test
    @DisplayName("excel writer test")
    void test_excel_writer() throws IOException {
        SimpleExcel simpleExcel = new SimpleExcel();
        SXSSFWorkbook workbook = simpleExcel.writeObjectsToWorkbook(SampleDto.class, list, "test");
        FileOutputStream fileInputStream = new FileOutputStream("C:\\Users\\mosic\\Downloads\\test.xlsx");
        workbook.write(fileInputStream);
    }

    @Test
    @DisplayName("List Set메서드 테스트")
    void test_list_set() {
//        Comparator<Integer> comparator = (s1, s2)->s2.compareTo(s1); // 내림차순
//        Comparator<Integer> comparator = (s1, s2)->s1.compareTo(s2); // 오름차순
        Comparator<Integer> comparator = Integer::compareTo; // 오름차순
        Map<Integer, Field> headerOrderMap = new TreeMap<Integer, Field>(comparator);

        ReflectionUtils.doWithFields(SampleDto.class, f -> {
            final ExcelColumn excelColumn = f.getAnnotation(ExcelColumn.class);
            if(excelColumn != null) {
                headerOrderMap.put(excelColumn.headerOrder(), f);
            }
        });

        list.forEach(sampleDto -> {
            headerOrderMap.forEach((integer, field) -> {
                System.out.println(integer + " / " + field);

                try {
                    ReflectionUtils.makeAccessible(field);
                    System.out.println(">> get: " +  field.get(sampleDto));
//                    Object o = (field.getType()) field.get(sampleDto));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            });
        });
        //list.forEach(System.out::println);
    }
}