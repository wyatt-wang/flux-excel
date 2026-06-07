package com.github.excel.pipeline.core;

public interface PipelineStep<C> {

	void execute(C context) throws Exception;

	default boolean cleanup() {
		return false;
	}

	default String name() {
		return getClass().getSimpleName();
	}
}
