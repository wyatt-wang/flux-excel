package com.github.excel.read.pipeline;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.ExcelReadKernel;

class ApplyCustomReaderStep<T extends ExcelBaseModel> implements ExcelReadStep<T> {

	private final ExcelReadKernel<T> kernel;

	ApplyCustomReaderStep(ExcelReadKernel<T> kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelReadContext<T> context) {
		kernel.applyCustomReader(context);
	}
}
