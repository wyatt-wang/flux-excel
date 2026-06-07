package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出excel异常
 */
public class ExcelWriterException extends RuntimeException implements Supplier<ExcelWriterException> {
	public ExcelWriterException(String message) {
		super(message);
	}

	public ExcelWriterException(Throwable cause) {
		super(cause);
	}

	@Override
	public ExcelWriterException get() {
		return this;
	}
}
