package com.github.excel.read.facade;

import com.github.excel.model.ExcelBaseModel;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 读取批处理
 */
public interface ExcelReaderBatchProcess<T extends ExcelBaseModel> {
	/**
	 * 获取批次大小
	 *
	 * @return
	 */
	int getBatchSize();

	/**
	 * 批处理
	 *
	 * @param dataList
	 */
	void process(List<T> dataList);

	/**
	 * 执行批处理
	 *
	 * @param dataList
	 */
	default void doProcess(List<T> dataList) {
		if (dataList.size() >= getBatchSize()) {
			process(dataList);
			dataList.clear();
		}
	}
}
