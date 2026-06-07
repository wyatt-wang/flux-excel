package com.github.excel.read.pipeline.event;

import com.github.excel.exception.ExcelReaderException;
import com.github.excel.read.handler.event.impl.ExcelEventXlsParseHandler;
import com.github.excel.read.handler.event.impl.ExcelEventXlsxParseHandler;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Locale;
import java.util.Objects;

public final class ExcelEventReadPipelines {

	private ExcelEventReadPipelines() {
	}

	public static <T> ExcelEventReadPipeline<T> eventReadPipeline() {
		return new ExcelEventReadPipeline<>(
				new ValidateEventReadStep<>(),
				new CreateEventReaderStep<>(),
				new ConfigureEventReaderStep<>(),
				new ProcessEventReadStep<>()
		);
	}

	private static class ValidateEventReadStep<T> implements ExcelEventReadStep<T> {
		@Override
		public void execute(ExcelEventReadContext<T> context) {
			if (context.getInputStream() == null && context.getFile() == null) {
				throw new ExcelReaderException("excel.input.is.empty");
			}
			if (Objects.isNull(context.getRowReader())) {
				throw new ExcelReaderException("excel.event.row.reader.is.empty");
			}
			if (Objects.isNull(context.getBatchHandler())) {
				throw new ExcelReaderException("excel.event.batch.handler.is.empty");
			}
			if (Objects.isNull(context.getFileName())) {
				throw new ExcelReaderException("excel.file.name.is.empty");
			}
		}
	}

	private static class CreateEventReaderStep<T> implements ExcelEventReadStep<T> {
		@Override
		public void execute(ExcelEventReadContext<T> context) {
			String lowerFileName = context.getFileName().toLowerCase(Locale.ROOT);
			if (lowerFileName.endsWith(".xlsx")) {
				context.setEventReader(new ExcelEventXlsxParseHandler<>());
				return;
			}
			if (lowerFileName.endsWith(".xls")) {
				context.setEventReader(new ExcelEventXlsParseHandler<>());
				return;
			}
			throw new ExcelReaderException("excel.file.type.not.support");
		}
	}

	private static class ConfigureEventReaderStep<T> implements ExcelEventReadStep<T> {
		@Override
		public void execute(ExcelEventReadContext<T> context) {
			context.getEventReader().setRowReader(context.getRowReader());
			context.getEventReader().setExecuteHandler(context.getBatchHandler());
		}
	}

	private static class ProcessEventReadStep<T> implements ExcelEventReadStep<T> {
		@Override
		public void execute(ExcelEventReadContext<T> context) throws Exception {
			if (context.getFile() != null) {
				try (InputStream stream = new FileInputStream(context.getFile())) {
					context.getEventReader().process(stream);
				}
				return;
			}
			context.getEventReader().process(context.getInputStream());
		}
	}
}
