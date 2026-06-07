package com.github.excel.read.pipeline.execution;

import java.util.Arrays;
import java.util.List;

public class ExcelReadExecutionPipeline {

	private final List<ExcelReadExecutionStep> steps;

	public ExcelReadExecutionPipeline(ExcelReadExecutionStep... steps) {
		this.steps = Arrays.asList(steps);
	}

	public void execute(ExcelReadExecutionContext context) {
		for (ExcelReadExecutionStep step : steps) {
			step.execute(context);
		}
	}
}
