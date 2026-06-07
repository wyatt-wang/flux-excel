package com.github.excel.pipeline.core;

import java.util.List;
import java.util.function.Function;

public final class LinearPipelineExecutor {

	public <C, E extends RuntimeException> void execute(List<? extends PipelineStep<C>> steps,
														C context,
														Function<Throwable, E> exceptionMapper) {
		Throwable failure = null;
		for (PipelineStep<C> step : steps) {
			if (step.cleanup()) {
				continue;
			}
			try {
				step.execute(context);
			} catch (Throwable e) {
				failure = e;
				break;
			}
		}
		Throwable cleanupFailure = runCleanup(steps, context);
		if (failure != null) {
			if (cleanupFailure != null) {
				failure.addSuppressed(cleanupFailure);
			}
			throw exceptionMapper.apply(failure);
		}
		if (cleanupFailure != null) {
			throw exceptionMapper.apply(cleanupFailure);
		}
	}

	private <C> Throwable runCleanup(List<? extends PipelineStep<C>> steps, C context) {
		Throwable failure = null;
		for (PipelineStep<C> step : steps) {
			if (!step.cleanup()) {
				continue;
			}
			try {
				step.execute(context);
				if (context instanceof CleanupAwarePipelineContext cleanupAwareContext) {
					cleanupAwareContext.setCleanupExecuted(true);
				}
			} catch (Throwable e) {
				if (failure == null) {
					failure = e;
				} else {
					failure.addSuppressed(e);
				}
			}
		}
		return failure;
	}
}
