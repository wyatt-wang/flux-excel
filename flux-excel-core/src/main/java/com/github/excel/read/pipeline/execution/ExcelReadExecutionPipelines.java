package com.github.excel.read.pipeline.execution;

public final class ExcelReadExecutionPipelines {

	private ExcelReadExecutionPipelines() {
	}

	public static ExcelReadExecutionPipeline fluentReadPipeline() {
		return new ExcelReadExecutionPipeline(
				new ParseReadStep(),
				new NotifyRowsStep(),
				new NotifyAfterAllStep()
		);
	}

	private static class ParseReadStep implements ExcelReadExecutionStep {
		@Override
		public void execute(ExcelReadExecutionContext context) {
			if (context.isCsvFile() && context.isCsvListRead()) {
				context.getCsvParser().run();
				return;
			}
			context.getWorkbookParser().run();
		}
	}

	private static class NotifyRowsStep implements ExcelReadExecutionStep {
		@Override
		public void execute(ExcelReadExecutionContext context) {
			context.getRowNotifier().run();
		}
	}

	private static class NotifyAfterAllStep implements ExcelReadExecutionStep {
		@Override
		public void execute(ExcelReadExecutionContext context) {
			context.getAfterAllNotifier().run();
		}
	}
}
