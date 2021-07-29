package com.github.donghyeok.excel.enums;

import lombok.Getter;

@Getter
public enum ExcelDataType {
	INTEGER(1),
	LONG(2),
	FLOAT(3),
	DOUBLE(4),
	DATE(5),
	STRING(6);

	private final int value;

	ExcelDataType(int value) {
		this.value = value;
	}

}
