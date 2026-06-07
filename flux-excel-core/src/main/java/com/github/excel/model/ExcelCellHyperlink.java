package com.github.excel.model;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Excel 单元格超链接信息
 */
@Data
@Accessors(chain = true)
public class ExcelCellHyperlink {
	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String cellRef;
	private String address;
	private String label;
	private String type;
	private String cellValue;
}
