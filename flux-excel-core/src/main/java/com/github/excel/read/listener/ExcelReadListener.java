package com.github.excel.read.listener;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReadError;

import java.util.Map;

/**
 * Excel 导入生命周期监听器
 */
public interface ExcelReadListener<T extends ExcelBaseModel> {
	default void onHeader(Map<Integer, Object> headMap, ExcelReadListenerContext context) {
	}

	default void onRow(T data, ExcelReadListenerContext context) {
	}

	default void onError(ExcelReadError error, ExcelReadListenerContext context) {
	}

	default boolean hasNext(ExcelReadListenerContext context) {
		return true;
	}

	default void afterAll(ExcelReadListenerContext context) {
	}
}
