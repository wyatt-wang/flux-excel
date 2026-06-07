package com.github.excel.write.pipeline.large;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.pipeline.core.LinearPipelineExecutor;
import com.github.excel.pipeline.core.PipelineStep;

import java.util.Arrays;
import java.util.ArrayList;
import java.util.List;

public class ExcelLargeWritePipeline {

	private final List<ExcelLargeWriteStep> steps;
	private final LinearPipelineExecutor executor = new LinearPipelineExecutor();

	public ExcelLargeWritePipeline(ExcelLargeWriteStep... steps) {
		this.steps = Arrays.asList(steps);
	}

	public void execute(ExcelLargeWriteContext context) {
		executor.execute(adaptSteps(), context, failure -> {
			if (failure instanceof ExcelWriterException) {
				return (ExcelWriterException) failure;
			}
			return new ExcelWriterException(failure);
		});
	}

	private List<PipelineStep<ExcelLargeWriteContext>> adaptSteps() {
		List<PipelineStep<ExcelLargeWriteContext>> adaptedSteps = new ArrayList<>(steps.size());
		for (ExcelLargeWriteStep step : steps) {
			adaptedSteps.add(new PipelineStep<ExcelLargeWriteContext>() {
				@Override
				public void execute(ExcelLargeWriteContext context) throws Exception {
					step.execute(context);
				}

				@Override
				public boolean cleanup() {
					return step.cleanup();
				}
			});
		}
		return adaptedSteps;
	}
}
