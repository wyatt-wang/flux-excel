package com.github.pipeline;

import com.github.excel.Excel;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.pipeline.ExcelReadContext;
import com.github.excel.read.pipeline.ExcelReadPipeline;
import com.github.excel.read.pipeline.ExcelReadStep;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

public class ExcelReadPipelineTest {

	@Test
	public void executesStepsInOrderThenCleanup() {
		List<String> order = new ArrayList<>();
		ExcelReadContext<ExcelBaseModel> context = ExcelReadContext.<ExcelBaseModel>builder().build();

		new ExcelReadPipeline<>(
				step("create", order),
				step("parse", order),
				step("validate", order),
				cleanup("cleanup", order)
		).execute(context);

		assertEquals("create", order.get(0));
		assertEquals("parse", order.get(1));
		assertEquals("validate", order.get(2));
		assertEquals("cleanup", order.get(3));
		assertTrue(context.isCleanupExecuted());
	}

	@Test
	public void executesCleanupWhenStepFails() {
		List<String> order = new ArrayList<>();
		ExcelReadContext<ExcelBaseModel> context = ExcelReadContext.<ExcelBaseModel>builder().build();

		try {
			new ExcelReadPipeline<>(
					step("create", order),
					failingStep("parse", order),
					step("after", order),
					cleanup("cleanup", order)
			).execute(context);
			fail("Pipeline should throw when a step fails");
		} catch (ExcelReaderException e) {
			assertEquals("parse failed", e.getMessage());
		}

		assertEquals("create", order.get(0));
		assertEquals("parse", order.get(1));
		assertEquals("cleanup", order.get(2));
		assertFalse(order.contains("after"));
		assertTrue(context.isCleanupExecuted());
	}

	@Test
	public void contextsAreIndependentBetweenExecutions() {
		ExcelReadContext<ExcelBaseModel> first = ExcelReadContext.<ExcelBaseModel>builder().build();
		ExcelReadContext<ExcelBaseModel> second = ExcelReadContext.<ExcelBaseModel>builder().build();
		List<ExcelReadContext<ExcelBaseModel>> contexts = new ArrayList<>();

		ExcelReadPipeline<ExcelBaseModel> pipeline = new ExcelReadPipeline<>(
				contexts::add,
				cleanup("cleanup", new ArrayList<>())
		);

		pipeline.execute(first);
		pipeline.execute(second);

		assertEquals(first, contexts.get(0));
		assertEquals(second, contexts.get(1));
		assertTrue(first.isCleanupExecuted());
		assertTrue(second.isCleanupExecuted());
	}

	@Test
	public void eventReadBuilderUsesPipelineValidation() {
		try {
			Excel.eventRead(new java.io.ByteArrayInputStream(new byte[0]))
					.fileName("events.xlsx")
					.parse();
			fail("Event read pipeline should validate row reader");
		} catch (ExcelReaderException e) {
			assertEquals("excel.event.row.reader.is.empty", e.getMessage());
		}
	}

	private ExcelReadStep<ExcelBaseModel> step(String name, List<String> order) {
		return context -> order.add(name);
	}

	private ExcelReadStep<ExcelBaseModel> failingStep(String name, List<String> order) {
		return context -> {
			order.add(name);
			throw new IllegalStateException(name + " failed");
		};
	}

	private ExcelReadStep<ExcelBaseModel> cleanup(String name, List<String> order) {
		return new ExcelReadStep<ExcelBaseModel>() {
			@Override
			public void execute(ExcelReadContext<ExcelBaseModel> context) {
				order.add(name);
			}

			@Override
			public boolean cleanup() {
				return true;
			}
		};
	}
}
