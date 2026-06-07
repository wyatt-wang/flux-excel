package com.github.excel.write;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 自定义写接口
 */
@FunctionalInterface
public interface ExcelCustomWriter {
	void execute(Workbook workbook);
}
