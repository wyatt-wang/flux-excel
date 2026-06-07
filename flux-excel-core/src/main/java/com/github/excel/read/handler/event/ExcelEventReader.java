package com.github.excel.read.handler.event;

import com.github.excel.exception.ExcelReaderException;
import com.github.excel.read.facade.AbstractEventBatchHandler;

import java.io.InputStream;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 事件模式下读取
 */
public interface ExcelEventReader<T> {

	/**
	 * 设置行解析器
	 * @param rowReader
	 */
	void setRowReader(ExcelEventRowReader<T> rowReader);

	/**
	 * 设置执行处理器
	 * @param executeHandler
	 */
	void setExecuteHandler(AbstractEventBatchHandler<T> executeHandler);

	/**
	 * 解析操作
	 * @param inputStream
	 */
	void process(InputStream inputStream) throws ExcelReaderException;

}
