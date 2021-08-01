package com.github.donghyeok.excel.example;

import com.github.donghyeok.excel.annotation.ExcelColumn;
import lombok.*;

@Getter
@Setter
@NoArgsConstructor(access = AccessLevel.PROTECTED)
@ToString
public class SampleDto {
    @ExcelColumn(headerName = "번호", headerOrder = 1)
    Integer no;
    @ExcelColumn(headerName = "업체명", headerOrder = 4, width = 9000)
    String companyName;
    @ExcelColumn(headerName = "대표자명", headerOrder = 5)
    String ceoName;
    @ExcelColumn(headerName = "이메일", headerOrder = 2, width = 9000)
    String email;
    @ExcelColumn(headerName = "연락처", headerOrder = 3, width = 9000)
    String phone;
    String address;
    String remark;

    @Builder
    public SampleDto(Integer no, String companyName, String ceoName, String email, String phone) {
        this.no = no;
        this.companyName = companyName;
        this.ceoName = ceoName;
        this.email = email;
        this.phone = phone;
    }
}
