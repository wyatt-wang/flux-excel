package com.github.excel.read.pipeline.event;

import com.github.excel.exception.ExcelReaderException;
import com.github.excel.pipeline.core.LinearPipelineExecutor;
import com.github.excel.pipeline.core.PipelineStep;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelEventReadPipeline<T> {

	private final List<ExcelEventReadStep<T>> steps;
	private final LinearPipelineExecutor executor = new LinearPipelineExecutor();

	@SafeVarargs
	public ExcelEventReadPipeline(ExcelEventReadStep<T>... steps) {
		this.steps = Arrays.asList(steps);
	}

	public void execute(ExcelEventReadContext<T> context) {
		executor.execute(adaptSteps(), context, failure -> {
			if (failure instanceof ExcelReaderException) {
				return (ExcelReaderException) failure;
			}
			return new ExcelReaderException(failure.getMessage());
		});
	}

	private List<PipelineStep<ExcelEventReadContext<T>>> adaptSteps() {
		List<PipelineStep<ExcelEventReadContext<T>>> adaptedSteps = new ArrayList<>(steps.size());
		for (ExcelEventReadStep<T> step : steps) {
			adaptedSteps.add(new PipelineStep<ExcelEventReadContext<T>>() {
				@Override
				public void execute(ExcelEventReadContext<T> context) throws Exception {
					step.execute(context);
				}
			});
		}
		return adaptedSteps;
	}
}
