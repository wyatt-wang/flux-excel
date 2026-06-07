package com.github.excel.model;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Excel 单元格批注信息
 */
@Data
@Accessors(chain = true)
public class ExcelCellComment {
	private Integer sheetIndex;
	private String sheetName;
	private Integer rowIndex;
	private Integer colIndex;
	private String cellRef;
	private String author;
	private String text;
	private String cellValue;
}
