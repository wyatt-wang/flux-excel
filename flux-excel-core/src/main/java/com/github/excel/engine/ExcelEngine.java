package com.github.excel.engine;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.ExcelReadKernel;
import com.github.excel.read.pipeline.ExcelReadPipeline;
import com.github.excel.read.pipeline.ExcelReadPipelines;
import com.github.excel.write.ExcelWriteKernel;
import com.github.excel.write.pipeline.ExcelWritePipeline;
import com.github.excel.write.pipeline.ExcelWritePipelines;

public final class ExcelEngine {

	private static final ExcelEngine DEFAULT = new ExcelEngine(ExcelRuntimeOptions.defaults());

	private final ExcelRuntimeOptions runtimeOptions;

	public ExcelEngine(ExcelRuntimeOptions runtimeOptions) {
		this.runtimeOptions = runtimeOptions;
	}

	public static ExcelEngine getDefault() {
		return DEFAULT;
	}

	public ExcelRuntimeOptions getRuntimeOptions() {
		return runtimeOptions;
	}

	public <T extends ExcelBaseModel> ExcelReadKernel<T> createReadKernel() {
		return new ExcelReadKernel<>(runtimeOptions);
	}

	public ExcelWriteKernel createWriteKernel() {
		return new ExcelWriteKernel();
	}

	public <T extends ExcelBaseModel> ExcelReadPipeline<T> createReadPipeline() {
		return ExcelReadPipelines.userReadPipeline(createReadKernel());
	}

	public ExcelWritePipeline createWritePipeline() {
		return ExcelWritePipelines.workbookPipeline(createWriteKernel());
	}
}
