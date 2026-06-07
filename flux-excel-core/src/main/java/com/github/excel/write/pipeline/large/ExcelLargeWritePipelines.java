package com.github.excel.write.pipeline.large;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.style.AbstractExcelStyle;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Objects;

public final class ExcelLargeWritePipelines {

	private ExcelLargeWritePipelines() {
	}

	public static ExcelLargeWritePipeline largeListPipeline() {
		return new ExcelLargeWritePipeline(
				new ApplyLargeStylesStep(),
				new WriteLargeListsStep(),
				new ExportLargeWorkbookStep(),
				new CloseLargeWorkbookStep()
		);
	}

	private static class ApplyLargeStylesStep implements ExcelLargeWriteStep {
		@Override
		public void execute(ExcelLargeWriteContext context) {
			for (Class<? extends AbstractExcelStyle> style : context.getStyles()) {
				context.getWriter().addStyle(style);
			}
		}
	}

	private static class WriteLargeListsStep implements ExcelLargeWriteStep {
		@Override
		public void execute(ExcelLargeWriteContext context) {
			ExcelLargeListWriter writer = context.getWriter();
			for (ExcelLargeWriteContext.ListOperation operation : context.getOperations()) {
				if (Objects.nonNull(operation.getExcludeFields())) {
					writer.process((java.util.List<ExcelBaseModel>) operation.getModelList(), operation.getExcludeFields(), operation.getModelClass());
				} else {
					writer.process((java.util.List<ExcelBaseModel>) operation.getModelList(), operation.getModelClass());
				}
			}
		}
	}

	private static class ExportLargeWorkbookStep implements ExcelLargeWriteStep {
		@Override
		public void execute(ExcelLargeWriteContext context) throws Exception {
			if (context.getFile() == null) {
				context.getWriter().export(context.getOutputStream());
				return;
			}
			try (OutputStream stream = new FileOutputStream(context.getFile())) {
				context.getWriter().export(stream);
			}
		}
	}

	private static class CloseLargeWorkbookStep implements ExcelLargeWriteStep {
		@Override
		public void execute(ExcelLargeWriteContext context) {
			if (context.getWriter() != null) {
				context.getWriter().close();
			}
		}

		@Override
		public boolean cleanup() {
			return true;
		}
	}
}
