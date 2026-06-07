package com.github.excel.read.pipeline;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.ExcelReadKernel;

class ValidateReadModelsStep<T extends ExcelBaseModel> implements ExcelReadStep<T> {

	private final ExcelReadKernel<T> kernel;

	ValidateReadModelsStep(ExcelReadKernel<T> kernel) {
		this.kernel = kernel;
	}

	@Override
	public void execute(ExcelReadContext<T> context) {
		kernel.validateModels(context);
	}
}
