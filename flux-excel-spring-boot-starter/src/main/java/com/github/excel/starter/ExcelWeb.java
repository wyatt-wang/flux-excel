package com.github.excel.starter;

import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.fluent.ExcelWriteBuilder;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import java.io.IOException;

/**
 * Servlet-specific fluent entry point kept out of flux-excel-core.
 */
public final class ExcelWeb {

	private ExcelWeb() {
	}

	public static ExcelWriteBuilder write(HttpServletRequest request, HttpServletResponse response, String fileName) {
		return write(request, response, fileName, ExcelSuffixEnum.XLSX);
	}

	public static ExcelWriteBuilder write(HttpServletRequest request, HttpServletResponse response,
										  String fileName, ExcelSuffixEnum suffixEnum) {
		try {
			ExcelWebUtil.setResponseHeader(request, response, fileName, suffixEnum.getSuffix());
			return ExcelWriteBuilder.toStream(response.getOutputStream())
					.fileName(fileName)
					.suffix(suffixEnum);
		} catch (IOException e) {
			throw new ExcelWriterException(e);
		}
	}
}
