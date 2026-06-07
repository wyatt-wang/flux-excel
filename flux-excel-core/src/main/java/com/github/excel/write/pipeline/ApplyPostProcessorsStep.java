package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class ApplyPostProcessorsStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	ApplyPostProcessorsStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.applyPostProcessors(context);
	}
}
