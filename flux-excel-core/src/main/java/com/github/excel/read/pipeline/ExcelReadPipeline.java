package com.github.excel.read.pipeline;

import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.pipeline.core.LinearPipelineExecutor;
import com.github.excel.pipeline.core.PipelineStep;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExcelReadPipeline<T extends ExcelBaseModel> {

	private final List<ExcelReadStep<T>> steps;
	private final LinearPipelineExecutor executor = new LinearPipelineExecutor();

	public ExcelReadPipeline(List<ExcelReadStep<T>> steps) {
		this.steps = new ArrayList<>(steps);
	}

	@SafeVarargs
	public ExcelReadPipeline(ExcelReadStep<T>... steps) {
		this(Arrays.asList(steps));
	}

	public void execute(ExcelReadContext<T> context) {
		executor.execute(adaptSteps(), context, failure -> {
			if (failure instanceof ExcelReaderException) {
				return (ExcelReaderException) failure;
			}
			return new ExcelReaderException(failure.getMessage());
		});
	}

	private List<PipelineStep<ExcelReadContext<T>>> adaptSteps() {
		List<PipelineStep<ExcelReadContext<T>>> adaptedSteps = new ArrayList<>(steps.size());
		for (ExcelReadStep<T> step : steps) {
			adaptedSteps.add(new PipelineStep<ExcelReadContext<T>>() {
				@Override
				public void execute(ExcelReadContext<T> context) throws Exception {
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
