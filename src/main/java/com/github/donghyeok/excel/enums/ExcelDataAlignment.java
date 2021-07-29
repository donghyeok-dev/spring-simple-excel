package com.github.donghyeok.excel.enums;

import lombok.Getter;

@Getter
public enum ExcelDataAlignment {
	LEFT(1),
	RIGHT(2),
	CENTER(3);

	private final int value;

	ExcelDataAlignment(int value) {
		this.value = value;
	}

}
