package com.github.donghyeok.excel.example;

import com.github.donghyeok.excel.annotation.SimpleExcelColumn;
import lombok.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

@Getter
@Setter
@NoArgsConstructor(access = AccessLevel.PROTECTED)
@ToString
public class SampleDto {
    @SimpleExcelColumn(headerName = "번호", columnOrder = 1, headerFontSize = 17,
            headerBackgroundColor = IndexedColors.BRIGHT_GREEN)
    Integer no;
    @SimpleExcelColumn(headerName = "업체명", columnOrder = 4, columnWidth = 9000,
            bodyFontSize = 15, bodyAlignment = HorizontalAlignment.LEFT, bodyBackgroundColor = IndexedColors.AQUA)
    String companyName;
    @SimpleExcelColumn(headerName = "대표자명", columnOrder = 5)
    String ceoName;
    @SimpleExcelColumn(headerName = "이메일", columnOrder = 2, columnWidth = 9000)
    String email;
    @SimpleExcelColumn(headerName = "연락처", columnOrder = 3, columnWidth = 9000)
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
