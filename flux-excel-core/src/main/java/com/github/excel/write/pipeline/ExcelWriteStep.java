package com.github.excel.write.pipeline;

public interface ExcelWriteStep {

	void execute(ExcelWriteContext context) throws Exception;

	default boolean cleanup() {
		return false;
	}

	default String name() {
		return getClass().getSimpleName();
	}
}
