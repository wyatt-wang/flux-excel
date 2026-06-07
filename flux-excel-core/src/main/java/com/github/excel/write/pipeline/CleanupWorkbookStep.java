package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class CleanupWorkbookStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	CleanupWorkbookStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.cleanup(context);
	}

	@Override
	public boolean cleanup() {
		return true;
	}
}
