package com.github.excel.write.pipeline.batch;

interface ExcelBatchWriteStep {

	void execute(ExcelBatchWriteContext context) throws Exception;
}
