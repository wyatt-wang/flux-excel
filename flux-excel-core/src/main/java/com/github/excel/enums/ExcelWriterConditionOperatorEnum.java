package com.github.excel.enums;

import org.apache.poi.ss.usermodel.ComparisonOperator;

/**
 * Excel 条件样式比较操作符
 */
public enum ExcelWriterConditionOperatorEnum {
	NO_COMPARISON(ComparisonOperator.NO_COMPARISON),
	BETWEEN(ComparisonOperator.BETWEEN),
	NOT_BETWEEN(ComparisonOperator.NOT_BETWEEN),
	EQUAL(ComparisonOperator.EQUAL),
	NOT_EQUAL(ComparisonOperator.NOT_EQUAL),
	GT(ComparisonOperator.GT),
	LT(ComparisonOperator.LT),
	GE(ComparisonOperator.GE),
	LE(ComparisonOperator.LE);

	private final byte poiCode;

	ExcelWriterConditionOperatorEnum(byte poiCode) {
		this.poiCode = poiCode;
	}

	public byte getPoiCode() {
		return poiCode;
	}
}
