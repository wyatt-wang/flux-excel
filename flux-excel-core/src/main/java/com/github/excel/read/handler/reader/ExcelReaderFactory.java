package com.github.excel.read.handler.reader;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.param.ExcelReaderFileParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.read.handler.event.ExcelEventReader;
import com.github.excel.read.handler.event.impl.ExcelEventXlsParseHandler;
import com.github.excel.read.handler.event.impl.ExcelEventXlsxParseHandler;

import java.util.Locale;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 读取器创建工厂
 */
public class ExcelReaderFactory {

	/**
	 * 创建用户模式读取
	 * @param readerStreamParam 流读取参数
	 * @return
	 */
	public static ExcelReader createUserReader(ExcelReaderStreamParam readerStreamParam) {
		return new ExcelReaderImpl(readerStreamParam);
	}

	public static ExcelReader createUserReader(String fileName, java.io.InputStream inputStream, String password, boolean closeInputStream) {
		ExcelReaderStreamParam readerParam = ExcelReaderStreamParam.builder()
				.stream(inputStream)
				.password(password)
				.template(fileName)
				.closeInputStream(closeInputStream)
				.build();
		return createUserReader(readerParam);
	}

	/**
	 * 创建用户模式读取
	 * @param readerFileParam 文件读取
	 * @return
	 */
	public static ExcelReader createUserReader(ExcelReaderFileParam readerFileParam) {
		return new ExcelReaderImpl(readerFileParam);
	}

	public static <T> ExcelEventReader<T> createExcelEventReader(String fileName) {
		String lowerFileName = fileName.toLowerCase(Locale.ROOT);
		if (lowerFileName.endsWith(ExcelConstant.XLSX_STR)) {
			return new ExcelEventXlsxParseHandler<>();
		}
		if (lowerFileName.endsWith(ExcelConstant.XLS_STR)) {
			return new ExcelEventXlsParseHandler<>();
		}
		throw new ExcelReaderException("excel.file.type.not.support");
	}

}
