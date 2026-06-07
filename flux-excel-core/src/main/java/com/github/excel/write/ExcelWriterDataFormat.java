package com.github.excel.write;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出格式化
 */
@FunctionalInterface
public interface ExcelWriterDataFormat {
	Object format(Object data, String pattern);
}
