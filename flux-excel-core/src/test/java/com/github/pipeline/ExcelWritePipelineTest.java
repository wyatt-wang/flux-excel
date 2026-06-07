package com.github.pipeline;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelWriteKernel;
import com.github.excel.write.pipeline.ExcelWriteContext;
import com.github.excel.write.pipeline.ExcelWritePipeline;
import com.github.excel.write.pipeline.ExcelWriteStep;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

public class ExcelWritePipelineTest {

	@Test
	public void executesStepsInOrderThenCleanup() {
		List<String> order = new ArrayList<>();
		ExcelWriteContext context = ExcelWriteContext.builder().build();

		new ExcelWritePipeline(
				step("create", order),
				step("styles", order),
				step("write", order),
				cleanup("cleanup", order)
		).execute(context);

		assertEquals("create", order.get(0));
		assertEquals("styles", order.get(1));
		assertEquals("write", order.get(2));
		assertEquals("cleanup", order.get(3));
		assertTrue(context.isCleanupExecuted());
	}

	@Test
	public void executesCleanupWhenStepFails() {
		List<String> order = new ArrayList<>();
		ExcelWriteContext context = ExcelWriteContext.builder().build();

		try {
			new ExcelWritePipeline(
					step("create", order),
					failingStep("write", order),
					step("after", order),
					cleanup("cleanup", order)
			).execute(context);
			fail("Pipeline should throw when a step fails");
		} catch (ExcelWriterException e) {
			assertEquals("write failed", e.getMessage());
		}

		assertEquals("create", order.get(0));
		assertEquals("write", order.get(1));
		assertEquals("cleanup", order.get(2));
		assertFalse(order.contains("after"));
		assertTrue(context.isCleanupExecuted());
	}

	@Test
	public void kernelDoesNotExtendLegacyBaseWriter() {
		assertFalse(BaseExcelWriter.class.isAssignableFrom(ExcelWriteKernel.class));
	}

	@Test
	public void contextsAreIndependentBetweenExecutions() {
		ExcelWriteContext first = ExcelWriteContext.builder().excelName("first.xlsx").build();
		ExcelWriteContext second = ExcelWriteContext.builder().excelName("second.xlsx").build();
		List<String> names = new ArrayList<>();

		ExcelWritePipeline pipeline = new ExcelWritePipeline(
				context -> names.add(context.getExcelName()),
				cleanup("cleanup", new ArrayList<>())
		);

		pipeline.execute(first);
		pipeline.execute(second);

		assertEquals("first.xlsx", names.get(0));
		assertEquals("second.xlsx", names.get(1));
		assertTrue(first.isCleanupExecuted());
		assertTrue(second.isCleanupExecuted());
	}

	private ExcelWriteStep step(String name, List<String> order) {
		return context -> order.add(name);
	}

	private ExcelWriteStep failingStep(String name, List<String> order) {
		return context -> {
			order.add(name);
			throw new IllegalStateException(name + " failed");
		};
	}

	private ExcelWriteStep cleanup(String name, List<String> order) {
		return new ExcelWriteStep() {
			@Override
			public void execute(ExcelWriteContext context) {
				order.add(name);
			}

			@Override
			public boolean cleanup() {
				return true;
			}
		};
	}
}
