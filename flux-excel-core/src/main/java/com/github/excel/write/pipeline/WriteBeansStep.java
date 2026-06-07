package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class WriteBeansStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	WriteBeansStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) throws Exception {
		kernel.writeBeans(context);
	}
}
