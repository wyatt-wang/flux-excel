package com.github.excel.read.handler.parser;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.engine.ExcelEngine;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.pipeline.ExcelReadContext;
import lombok.extern.slf4j.Slf4j;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 用户模式下解析bean处理器
 */
@Slf4j
public class ExcelUserReaderParserImpl<T extends ExcelBaseModel> implements ExcelReaderUserParser {

	private final ExcelReaderContext<T> readerContext;

	public ExcelUserReaderParserImpl(ExcelReaderContext<T> readerContext) {
		this.readerContext = readerContext;
	}


	@Override
	public void process() {
		ExcelReadContext<T> context = ExcelReadContext.<T>builder()
				.readerContext(readerContext)
				.build();
		ExcelEngine.getDefault().<T>createReadPipeline().execute(context);
	}
}
