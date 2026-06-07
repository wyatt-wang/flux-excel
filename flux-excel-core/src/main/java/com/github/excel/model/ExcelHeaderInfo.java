package com.github.excel.model;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Excel 表头信息
 */
@Data
@Accessors(chain = true)
public class ExcelHeaderInfo {
	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String title;
}
