package com.github.excel.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 单元格样式配置器
 */
@FunctionalInterface
public interface ExcelCellStyleConfigurer {
	void configure(CellStyle cellStyle, Workbook workbook);
}
