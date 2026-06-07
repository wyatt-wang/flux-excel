package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class WriteOutputStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	WriteOutputStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) throws Exception {
		kernel.writeOutput(context);
	}
}
