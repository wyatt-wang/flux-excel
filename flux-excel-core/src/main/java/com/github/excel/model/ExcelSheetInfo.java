package com.github.excel.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * Excel sheet 元信息
 */
@Data
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class ExcelSheetInfo {
	private int index;
	private String name;
	private int firstRowNum;
	private int lastRowNum;
	private int physicalNumberOfRows;
}
