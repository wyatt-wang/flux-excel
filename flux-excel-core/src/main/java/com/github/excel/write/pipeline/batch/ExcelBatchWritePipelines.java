package com.github.excel.write.pipeline.batch;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.AbstractExcelStyle;

import java.util.Objects;

public final class ExcelBatchWritePipelines {

	private ExcelBatchWritePipelines() {
	}

	public static ExcelBatchWritePipeline zipPipeline() {
		return new ExcelBatchWritePipeline(
				new ApplyBatchStylesStep(),
				new WriteBatchListsStep(),
				context -> context.getWriter().export(context.getZipFileName())
		);
	}

	public static ExcelBatchWritePipeline responsePipeline() {
		return new ExcelBatchWritePipeline(
				new ApplyBatchStylesStep(),
				new WriteBatchListsStep(),
				context -> context.getWriter().export(context.getRequest(), context.getResponse(), context.getResponseFileName())
		);
	}

	private static class ApplyBatchStylesStep implements ExcelBatchWriteStep {
		@Override
		public void execute(ExcelBatchWriteContext context) {
			for (Class<? extends AbstractExcelStyle> style : context.getStyles()) {
				context.getWriter().addStyle(style);
			}
		}
	}

	private static class WriteBatchListsStep implements ExcelBatchWriteStep {
		@Override
		public void execute(ExcelBatchWriteContext context) {
			for (ExcelBatchWriteContext.ListOperation operation : context.getOperations()) {
				if (Objects.nonNull(operation.getExcludeFields())) {
					context.getWriter().process((java.util.List<ExcelBaseModel>) operation.getModelList(),
							operation.getFileName(), operation.getSheetName(), operation.getExcludeFields(), operation.getModelClass());
				} else {
					context.getWriter().process((java.util.List<ExcelBaseModel>) operation.getModelList(),
							operation.getFileName(), operation.getSheetName(), operation.getModelClass());
				}
			}
		}
	}
}
