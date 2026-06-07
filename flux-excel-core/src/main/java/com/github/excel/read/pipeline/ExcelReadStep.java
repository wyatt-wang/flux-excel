package com.github.excel.read.pipeline;

import com.github.excel.model.ExcelBaseModel;

public interface ExcelReadStep<T extends ExcelBaseModel> {

	void execute(ExcelReadContext<T> context) throws Exception;

	default boolean cleanup() {
		return false;
	}

	default String name() {
		return getClass().getSimpleName();
	}
}
