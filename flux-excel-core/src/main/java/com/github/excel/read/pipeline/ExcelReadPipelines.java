package com.github.excel.read.pipeline;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.ExcelReadKernel;

public final class ExcelReadPipelines {

	private ExcelReadPipelines() {
	}

	public static <T extends ExcelBaseModel> ExcelReadPipeline<T> userReadPipeline(ExcelReadKernel<T> kernel) {
		return new ExcelReadPipeline<>(
				new CreateReadWorkbookStep<>(kernel),
				new PrepareReadRuntimeStep<>(kernel),
				new ValidateReadTemplateStep<>(kernel),
				new ParseReadSheetsStep<>(kernel),
				new ApplyCustomReaderStep<>(kernel),
				new ValidateReadModelsStep<>(kernel),
				new FlushReadBatchStep<>(kernel),
				new ClearReadCoordinatesStep<>(kernel),
				new CleanupReadWorkbookStep<>(kernel)
		);
	}
}
