package com.github.excel.write.pipeline.large;

interface ExcelLargeWriteStep {

	void execute(ExcelLargeWriteContext context) throws Exception;

	default boolean cleanup() {
		return false;
	}
}
