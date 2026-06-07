package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 读取excel异常
 */
public class ExcelReaderException extends RuntimeException implements Supplier<ExcelReaderException> {
	public ExcelReaderException(String message) {
		super(message);
	}

	@Override
	public ExcelReaderException get() {
		return this;
	}
}
