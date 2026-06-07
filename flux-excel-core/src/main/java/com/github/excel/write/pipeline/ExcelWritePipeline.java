package com.github.excel.write.pipeline;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.pipeline.core.LinearPipelineExecutor;
import com.github.excel.pipeline.core.PipelineStep;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelWritePipeline {

	private final List<ExcelWriteStep> steps;
	private final LinearPipelineExecutor executor = new LinearPipelineExecutor();

	public ExcelWritePipeline(List<ExcelWriteStep> steps) {
		this.steps = new ArrayList<>(steps);
	}

	public ExcelWritePipeline(ExcelWriteStep... steps) {
		this(Arrays.asList(steps));
	}

	public void execute(ExcelWriteContext context) {
		executor.execute(adaptSteps(), context, failure -> {
			if (failure instanceof ExcelWriterException) {
				return (ExcelWriterException) failure;
			}
			return new ExcelWriterException(failure.getMessage());
		});
	}

	private List<PipelineStep<ExcelWriteContext>> adaptSteps() {
		List<PipelineStep<ExcelWriteContext>> adaptedSteps = new ArrayList<>(steps.size());
		for (ExcelWriteStep step : steps) {
			adaptedSteps.add(new PipelineStep<ExcelWriteContext>() {
				@Override
				public void execute(ExcelWriteContext context) throws Exception {
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
