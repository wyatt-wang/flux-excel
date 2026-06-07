package com.github.excel.fluent;

import com.github.excel.read.handler.event.ExcelEventRowReader;
import com.github.excel.read.facade.AbstractEventBatchHandler;
import com.github.excel.read.pipeline.event.ExcelEventReadContext;
import com.github.excel.read.pipeline.event.ExcelEventReadPipelines;

import java.io.File;
import java.io.InputStream;

public class ExcelEventReadBuilder<T> {

	private final InputStream inputStream;
	private final File file;
	private String fileName;
	private ExcelEventRowReader<T> rowReader;
	private AbstractEventBatchHandler<T> batchHandler;

	private ExcelEventReadBuilder(InputStream inputStream, File file) {
		this.inputStream = inputStream;
		this.file = file;
		if (file != null) {
			this.fileName = file.getName();
		}
	}


	public static <T> ExcelEventReadBuilder<T> fromFile(File file) {
		return new ExcelEventReadBuilder<>(null, file);
	}

	public static <T> ExcelEventReadBuilder<T> fromStream(InputStream inputStream) {
		return new ExcelEventReadBuilder<>(inputStream, null);
	}

	public ExcelEventReadBuilder<T> fileName(String fileName) {
		this.fileName = fileName;
		return this;
	}

	public ExcelEventReadBuilder<T> rowReader(ExcelEventRowReader<T> rowReader) {
		this.rowReader = rowReader;
		return this;
	}

	public ExcelEventReadBuilder<T> batchHandler(AbstractEventBatchHandler<T> batchHandler) {
		this.batchHandler = batchHandler;
		return this;
	}

	public ExcelEventReadBuilder<T> parse() {
		ExcelEventReadContext<T> context = new ExcelEventReadContext<>();
		context.setInputStream(inputStream);
		context.setFile(file);
		context.setFileName(fileName);
		context.setRowReader(rowReader);
		context.setBatchHandler(batchHandler);
		ExcelEventReadPipelines.<T>eventReadPipeline().execute(context);
		return this;
	}

}
