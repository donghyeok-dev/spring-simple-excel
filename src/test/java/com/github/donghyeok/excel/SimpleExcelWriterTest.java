package com.github.donghyeok.excel;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import com.github.donghyeok.excel.example.SampleDto;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.util.ReflectionUtils;
import org.springframework.util.SocketUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.util.*;
import static org.junit.jupiter.api.Assertions.*;

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
    @DisplayName("숫자형 체크")
    void test_numberic() {
        assertTrue(isNumberic(int.class));
        assertTrue(isNumberic(Double.class));
        assertTrue(isNumberic(Float.class));
        assertFalse(isNumberic(String.class));
        assertFalse(isNumberic(BigInteger.class));
    }

    @FunctionalInterface
    interface Calculation<T> {
        T apply(T x, T y);
    }

    @Test
    @DisplayName("object sum")
    void test_object_sum() {
        Calculation addition;

        Class<?> tClass = Integer.class;
        Object sum=0;
        Calculation<Integer> integerAddition = Integer::sum;
        addition = integerAddition;

        for(int i=0; i<10; i++) {
            //sum = tClass.cast(a);
            sum = addition.apply(sum, 10);
        }
        System.out.println(sum);

        sum = 0f;
        Calculation<Float> addition2 = Float::sum;
        addition = addition2;

        for(int i=0; i<10; i++) {
            //sum = tClass.cast(a);
            sum = addition.apply(sum, 10.1f);
        }
        System.out.println(sum);
    }

    protected boolean isNumberic(Class<?> tClass) {
        return (Integer.class.equals(tClass)
                || int.class.equals(tClass)
                || Double.class.equals(tClass)
                || double.class.equals(tClass)
                || Float.class.equals(tClass)
                || float.class.equals(tClass));
    }

    @Test
    @DisplayName("List Set메서드 테스트")
    void test_list_set() {
//        Comparator<Integer> comparator = (s1, s2)->s2.compareTo(s1); // 내림차순
//        Comparator<Integer> comparator = (s1, s2)->s1.compareTo(s2); // 오름차순
        Comparator<Integer> comparator = Integer::compareTo; // 오름차순
        Map<Integer, Field> columnOrderMap = new TreeMap<Integer, Field>(comparator);

        ReflectionUtils.doWithFields(SampleDto.class, f -> {
            final SimpleExcelColumn simpleExcelColumn = f.getAnnotation(SimpleExcelColumn.class);
            if(simpleExcelColumn != null) {
                columnOrderMap.put(simpleExcelColumn.columnOrder(), f);
            }
        });

        list.forEach(sampleDto -> {
            columnOrderMap.forEach((integer, field) -> {
                System.out.println(integer + " / " + field);
                try {
                    ReflectionUtils.makeAccessible(field);
                    System.out.println(">> get: " +  field.get(sampleDto));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            });
        });
    }
}