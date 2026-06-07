package com.github.excel.read.facade;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 自定义写读接口
 */
@FunctionalInterface
public interface ExcelCustomReader {
	void read(Workbook workbook);
}
