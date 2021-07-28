package com.github.donghyeok.excel.example;

import com.github.donghyeok.excel.annotation.ExcelColumn;
import lombok.*;

@Getter
@Setter
@NoArgsConstructor(access = AccessLevel.PROTECTED)
@ToString
public class SampleDto {
    @ExcelColumn(headerName = "번호")
    Integer no;
    @ExcelColumn(headerName = "업체명")
    String companyName;
    @ExcelColumn(headerName = "대표자명")
    String ceoName;
    @ExcelColumn(headerName = "이메일")
    String email;
    @ExcelColumn(headerName = "연락처")
    String phone;
    String address;
    String remark;
}
