package com.github.excel.enums;

import com.github.excel.constant.ExcelConstant;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel后缀
 */
public enum ExcelSuffixEnum {

	XLS(ExcelConstant.XLS_STR),

	XLSX(ExcelConstant.XLSX_STR),

	CSV(ExcelConstant.CSV_STR);

	ExcelSuffixEnum(String suffix){
		this.suffix = suffix;
	}

	private String suffix ;

	public String getSuffix() {
		return suffix;
	}
}
