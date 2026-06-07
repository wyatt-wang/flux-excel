package com.github.excel.write.pipeline.batch;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.pipeline.core.LinearPipelineExecutor;
import com.github.excel.pipeline.core.PipelineStep;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelBatchWritePipeline {

	private final List<ExcelBatchWriteStep> steps;
	private final LinearPipelineExecutor executor = new LinearPipelineExecutor();

	public ExcelBatchWritePipeline(ExcelBatchWriteStep... steps) {
		this.steps = Arrays.asList(steps);
	}

	public void execute(ExcelBatchWriteContext context) {
		executor.execute(adaptSteps(), context, failure -> {
			if (failure instanceof ExcelWriterException) {
				return (ExcelWriterException) failure;
			}
			return new ExcelWriterException(failure);
		});
	}

	private List<PipelineStep<ExcelBatchWriteContext>> adaptSteps() {
		List<PipelineStep<ExcelBatchWriteContext>> adaptedSteps = new ArrayList<>(steps.size());
		for (ExcelBatchWriteStep step : steps) {
			adaptedSteps.add(new PipelineStep<ExcelBatchWriteContext>() {
				@Override
				public void execute(ExcelBatchWriteContext context) throws Exception {
					step.execute(context);
				}
			});
		}
		return adaptedSteps;
	}
}
