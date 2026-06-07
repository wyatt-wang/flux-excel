package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class CreateWorkbookStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	CreateWorkbookStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.createWorkbook(context);
	}
}
