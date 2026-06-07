package com.github.excel.read.pipeline;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.ExcelReadKernel;

class CleanupReadWorkbookStep<T extends ExcelBaseModel> implements ExcelReadStep<T> {

	private final ExcelReadKernel<T> kernel;

	CleanupReadWorkbookStep(ExcelReadKernel<T> kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelReadContext<T> context) {
		kernel.cleanup(context);
	}

	@Override
	public boolean cleanup() {
		return true;
	}
}
