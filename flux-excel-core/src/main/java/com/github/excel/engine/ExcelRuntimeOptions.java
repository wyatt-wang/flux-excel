package com.github.excel.engine;

import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.read.executor.impl.ExcelReaderPictureStanderExecutor;
import com.github.excel.read.executor.impl.ExcelReaderTemplateStanderValidator;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.read.format.ExcelDefaultReaderDataFormat;
import com.github.excel.read.format.ExcelReaderDataFormat;
import com.github.excel.read.format.ExcelReaderFormatManager;
import com.github.excel.read.handler.row.ExcelReaderRowParser;
import com.github.excel.read.handler.row.ExcelReaderRowParserImpl;

import java.util.ArrayList;
import java.util.List;
import java.util.ServiceLoader;
import java.util.function.Supplier;

public final class ExcelRuntimeOptions {

	private final Supplier<ExcelReaderFormatManager> readerFormatManagerSupplier;
	private final Supplier<ExcelReaderTemplateValidator> templateValidatorSupplier;
	private final Supplier<ExcelReaderPictureExecutor> pictureExecutorSupplier;
	private final Supplier<ExcelReaderRowParser<?>> rowParserSupplier;
	private final Supplier<ExcelReaderDataFormat> csvDataFormatSupplier;

	private ExcelRuntimeOptions(Builder builder) {
		this.readerFormatManagerSupplier = builder.readerFormatManagerSupplier;
		this.templateValidatorSupplier = builder.templateValidatorSupplier;
		this.pictureExecutorSupplier = builder.pictureExecutorSupplier;
		this.rowParserSupplier = builder.rowParserSupplier;
		this.csvDataFormatSupplier = builder.csvDataFormatSupplier;
	}

	public static ExcelRuntimeOptions defaults() {
		return builder().build();
	}

	public static Builder builder() {
		return new Builder();
	}

	public ExcelReaderFormatManager createReaderFormatManager() {
		return readerFormatManagerSupplier.get();
	}

	public ExcelReaderTemplateValidator createTemplateValidator() {
		return templateValidatorSupplier.get();
	}

	public ExcelReaderPictureExecutor createPictureExecutor() {
		return pictureExecutorSupplier.get();
	}

	@SuppressWarnings("unchecked")
	public <T extends ExcelBaseModel> ExcelReaderRowParser<T> createRowParser() {
		return (ExcelReaderRowParser<T>) rowParserSupplier.get();
	}

	public ExcelReaderDataFormat createCsvDataFormat() {
		return csvDataFormatSupplier.get();
	}

	private static ExcelReaderRowParser<?> loadRowParser() {
		ServiceLoader<ExcelReaderRowParser> serviceLoader = ServiceLoader.load(ExcelReaderRowParser.class);
		List<ExcelReaderRowParser> rowParsers = new ArrayList<>();
		for (ExcelReaderRowParser rowParser : serviceLoader) {
			rowParsers.add(rowParser);
		}
		return rowParsers.stream()
				.filter(rowParser -> !(rowParser instanceof ExcelReaderRowParserImpl))
				.findFirst()
				.orElse(new ExcelReaderRowParserImpl<>());
	}

	public static final class Builder {
		private Supplier<ExcelReaderFormatManager> readerFormatManagerSupplier = ExcelReaderFormatManager::new;
		private Supplier<ExcelReaderTemplateValidator> templateValidatorSupplier = ExcelReaderTemplateStanderValidator::new;
		private Supplier<ExcelReaderPictureExecutor> pictureExecutorSupplier = ExcelReaderPictureStanderExecutor::new;
		private Supplier<ExcelReaderRowParser<?>> rowParserSupplier = ExcelRuntimeOptions::loadRowParser;
		private Supplier<ExcelReaderDataFormat> csvDataFormatSupplier = ExcelDefaultReaderDataFormat::new;

		private Builder() {
		}

		public Builder readerFormatManager(Supplier<ExcelReaderFormatManager> supplier) {
			this.readerFormatManagerSupplier = supplier;
			return this;
		}

		public Builder templateValidator(Supplier<ExcelReaderTemplateValidator> supplier) {
			this.templateValidatorSupplier = supplier;
			return this;
		}

		public Builder pictureExecutor(Supplier<ExcelReaderPictureExecutor> supplier) {
			this.pictureExecutorSupplier = supplier;
			return this;
		}

		public Builder rowParser(Supplier<ExcelReaderRowParser<?>> supplier) {
			this.rowParserSupplier = supplier;
			return this;
		}

		public Builder csvDataFormat(Supplier<ExcelReaderDataFormat> supplier) {
			this.csvDataFormatSupplier = supplier;
			return this;
		}

		public ExcelRuntimeOptions build() {
			return new ExcelRuntimeOptions(this);
		}
	}
}
