package com.github.excel.read.pipeline.event;

interface ExcelEventReadStep<T> {

	void execute(ExcelEventReadContext<T> context) throws Exception;
}
