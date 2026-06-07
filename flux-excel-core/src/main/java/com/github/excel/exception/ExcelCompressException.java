package com.github.excel.exception;

import java.util.function.Supplier;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 压缩文件异常
 */
public class ExcelCompressException extends RuntimeException implements Supplier<ExcelCompressException> {
	public ExcelCompressException(String message) {
		super(message);
	}

	@Override
	public ExcelCompressException get() {
		return this;
	}
}
