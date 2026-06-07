package com.github.excel.read.format;

import java.text.ParseException;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel读取格式化
 */
@FunctionalInterface
public interface ExcelReaderDataFormat {
	Object format(Object data, String pattern, Class<?> targetCla) throws ParseException;
}
