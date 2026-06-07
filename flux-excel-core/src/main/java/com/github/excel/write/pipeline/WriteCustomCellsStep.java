package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

class WriteCustomCellsStep implements ExcelWriteStep {

	private final ExcelWriteKernel kernel;

	WriteCustomCellsStep(ExcelWriteKernel kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelWriteContext context) {
		kernel.writeCustomCells(context);
	}
}
