package com.github.excel.read.handler.reader;

import com.github.excel.param.ExcelReaderFileParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.read.handler.parser.ExcelReaderUserParser;
import com.github.excel.read.handler.parser.ExcelUserReaderParserImpl;
import lombok.extern.slf4j.Slf4j;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 用户模式下读取
 */
@Slf4j
public class ExcelReaderImpl extends AbstractExcelReader {

	public ExcelReaderImpl(ExcelReaderStreamParam param) {
		super(param);
	}

	public ExcelReaderImpl(ExcelReaderFileParam param) {
		super(param);
	}

	@Override
	public ExcelReaderUserParser createHandler() {
		return new ExcelUserReaderParserImpl(readerContext);
	}
}
