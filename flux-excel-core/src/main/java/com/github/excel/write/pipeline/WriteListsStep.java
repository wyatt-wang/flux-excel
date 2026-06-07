package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class WriteListsStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	WriteListsStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.writeLists(context);
	}
}
