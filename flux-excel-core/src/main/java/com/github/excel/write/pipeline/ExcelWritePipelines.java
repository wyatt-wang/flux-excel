package com.github.excel.write.pipeline;

import com.github.excel.write.ExcelWriteKernel;

public final class ExcelWritePipelines {

	private ExcelWritePipelines() {
	}

	public static ExcelWritePipeline workbookPipeline(ExcelWriteKernel kernel) {
		return new ExcelWritePipeline(
				new CreateWorkbookStep(kernel),
				new PrepareStylesStep(kernel),
				new WriteCustomCellsStep(kernel),
				new WriteBeansStep(kernel),
				new WriteListsStep(kernel),
				new ApplyPostProcessorsStep(kernel),
				new WriteOutputStep(kernel),
				new CleanupWorkbookStep(kernel)
		);
	}
}
