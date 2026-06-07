package com.github.excel.model;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * Excel 合并单元格区域信息
 */
@Data
@Accessors(chain = true)
public class ExcelMergedCell {
	private Integer sheetIndex;
	private String sheetName;
	private Integer firstRow;
	private Integer lastRow;
	private Integer firstCol;
	private Integer lastCol;
	private String cellRange;
	private String value;
}
