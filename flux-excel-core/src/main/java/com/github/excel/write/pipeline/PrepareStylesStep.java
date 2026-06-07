package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class PrepareStylesStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	PrepareStylesStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.prepareStyles(context);
	}
}
