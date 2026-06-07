package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * Excel 导入错误明细
 */
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class ExcelReadError {
	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String fieldName;
	private String titleName;
	private Object rawValue;
	private String message;
}
