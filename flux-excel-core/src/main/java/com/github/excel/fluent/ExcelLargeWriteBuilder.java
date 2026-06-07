package com.github.excel.fluent;

import com.github.excel.exception.ExcelWriterException;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.ExcelLargeListWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.excel.write.pipeline.large.ExcelLargeWriteContext;
import com.github.excel.write.pipeline.large.ExcelLargeWritePipelines;
import com.github.excel.write.style.AbstractExcelStyle;

import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelLargeWriteBuilder {

	private final OutputStream outputStream;
	private final File file;
	private ExcelLargeListWriter writer;
	private String sheetName = "sheet";
	private int sheetRowMaxCount = 0;
	private int rowAccessWindowSize = ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE;
	private boolean compressTempFiles = false;
	private boolean useSharedStringsTable = false;
	private Boolean noneDataTips;
	private Class<? extends ExcelBaseModel> modelClass;
	private final List<Class<? extends AbstractExcelStyle>> styles = new ArrayList<>();
	private final List<ExcelLargeWriteContext.ListOperation> operations = new ArrayList<>();

	private ExcelLargeWriteBuilder(OutputStream outputStream, File file) {
		this.outputStream = outputStream;
		this.file = file;
	}

	public static ExcelLargeWriteBuilder toStream(OutputStream outputStream) {
		return new ExcelLargeWriteBuilder(outputStream, null);
	}

	public static ExcelLargeWriteBuilder toFile(File file) {
		return new ExcelLargeWriteBuilder(null, file);
	}

	public ExcelLargeWriteBuilder sheet(String sheetName) {
		this.sheetName = sheetName;
		if (writer != null) {
			writer.setSheetName(sheetName);
		}
		return this;
	}

	public ExcelLargeWriteBuilder sheetRowMaxCount(int sheetRowMaxCount) {
		this.sheetRowMaxCount = sheetRowMaxCount;
		return this;
	}

	public ExcelLargeWriteBuilder rowAccessWindowSize(int rowAccessWindowSize) {
		this.rowAccessWindowSize = rowAccessWindowSize;
		return this;
	}

	public ExcelLargeWriteBuilder compressTempFiles(boolean compressTempFiles) {
		this.compressTempFiles = compressTempFiles;
		return this;
	}

	public ExcelLargeWriteBuilder useSharedStringsTable(boolean useSharedStringsTable) {
		this.useSharedStringsTable = useSharedStringsTable;
		return this;
	}

	public ExcelLargeWriteBuilder modelClass(Class<? extends ExcelBaseModel> modelClass) {
		this.modelClass = modelClass;
		if (writer != null) {
			writer.setCla(modelClass);
		}
		return this;
	}

	public ExcelLargeWriteBuilder noneDataTips(boolean noneDataTips) {
		this.noneDataTips = noneDataTips;
		if (writer != null) {
			writer.setNoneDataTips(noneDataTips);
		}
		return this;
	}

	public ExcelLargeWriteBuilder style(Class<? extends AbstractExcelStyle> styleClass) {
		styles.add(styleClass);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelLargeWriteBuilder list(List<T> modelList) {
		if (modelClass == null) {
			throw new ExcelWriterException("modelClass required before list");
		}
		addListOperation(modelList, modelClass, null);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelLargeWriteBuilder list(List<T> modelList, Class<? extends ExcelBaseModel> modelClass) {
		this.modelClass = modelClass;
		addListOperation(modelList, modelClass, null);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelLargeWriteBuilder list(List<T> modelList, Class<? extends ExcelBaseModel> modelClass, String... excludeFields) {
		this.modelClass = modelClass;
		addListOperation(modelList, modelClass, excludeFields);
		return this;
	}

	public ExcelLargeWriteBuilder export() {
		ExcelLargeWriteContext context = new ExcelLargeWriteContext();
		context.setWriter(writer());
		context.setOutputStream(outputStream);
		context.setFile(file);
		context.setStyles(styles);
		context.setOperations(operations);
		ExcelLargeWritePipelines.largeListPipeline().execute(context);
		return this;
	}

	public void close() {
		if (writer != null) {
			writer.close();
		}
	}

	public ExcelLargeListWriter writer() {
		if (writer == null) {
			writer = createWriter();
		}
		return writer;
	}

	private ExcelLargeListWriter createWriter() {
		ExcelLargeListWriter largeListWriter;
		if (sheetRowMaxCount > 0 && modelClass != null) {
			largeListWriter = ExcelWriterFactory.createLargeListWriter(sheetName, sheetRowMaxCount, modelClass,
					rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
		} else if (sheetRowMaxCount > 0) {
			largeListWriter = ExcelWriterFactory.createLargeListWriter(sheetName, sheetRowMaxCount,
					rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
		} else {
			largeListWriter = ExcelWriterFactory.createLargeListWriter(sheetName, rowAccessWindowSize,
					compressTempFiles, useSharedStringsTable);
		}
		if (modelClass != null) {
			largeListWriter.setCla(modelClass);
		}
		if (noneDataTips != null) {
			largeListWriter.setNoneDataTips(noneDataTips);
		}
		return largeListWriter;
	}

	private <T extends ExcelBaseModel> void addListOperation(List<T> modelList, Class<? extends ExcelBaseModel> modelClass, String[] excludeFields) {
		ExcelLargeWriteContext.ListOperation operation = new ExcelLargeWriteContext.ListOperation();
		operation.setModelList(modelList);
		operation.setModelClass(modelClass);
		operation.setExcludeFields(excludeFields);
		operations.add(operation);
	}
}
