package com.github.excel.fluent;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.write.ExcelLargeListBatchWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.excel.write.pipeline.batch.ExcelBatchWriteContext;
import com.github.excel.write.pipeline.batch.ExcelBatchWritePipelines;
import com.github.excel.write.style.AbstractExcelStyle;

import java.util.ArrayList;
import java.util.List;

public class ExcelBatchWriteBuilder {

	private ExcelLargeListBatchWriter writer;
	private final String outputDirPath;
	private final Integer maxPoolSize;
	private String fileName = "excel";
	private String sheetName = "sheet";
	private int rowAccessWindowSize = ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE;
	private boolean compressTempFiles = false;
	private boolean useSharedStringsTable = false;
	private final List<Class<? extends AbstractExcelStyle>> styles = new ArrayList<>();
	private final List<ExcelBatchWriteContext.ListOperation> operations = new ArrayList<>();

	public ExcelBatchWriteBuilder(String outputDirPath) {
		this(outputDirPath, null);
	}

	public ExcelBatchWriteBuilder(String outputDirPath, int maxPoolSize) {
		this(outputDirPath, Integer.valueOf(maxPoolSize));
	}

	private ExcelBatchWriteBuilder(String outputDirPath, Integer maxPoolSize) {
		this.outputDirPath = outputDirPath;
		this.maxPoolSize = maxPoolSize;
	}

	public ExcelBatchWriteBuilder file(String fileName) {
		this.fileName = fileName;
		return this;
	}

	public ExcelBatchWriteBuilder sheet(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}

	public ExcelBatchWriteBuilder style(Class<? extends AbstractExcelStyle> styleClass) {
		styles.add(styleClass);
		return this;
	}

	public ExcelBatchWriteBuilder rowAccessWindowSize(int rowAccessWindowSize) {
		this.rowAccessWindowSize = rowAccessWindowSize;
		return this;
	}

	public ExcelBatchWriteBuilder compressTempFiles(boolean compressTempFiles) {
		this.compressTempFiles = compressTempFiles;
		return this;
	}

	public ExcelBatchWriteBuilder useSharedStringsTable(boolean useSharedStringsTable) {
		this.useSharedStringsTable = useSharedStringsTable;
		return this;
	}

	public <T extends ExcelBaseModel> ExcelBatchWriteBuilder list(List<T> modelList, Class<? extends ExcelBaseModel> modelClass) {
		addListOperation(modelList, modelClass, null);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelBatchWriteBuilder list(List<T> modelList, Class<? extends ExcelBaseModel> modelClass, String... excludeFields) {
		addListOperation(modelList, modelClass, excludeFields);
		return this;
	}

	public ExcelBatchWriteBuilder exportZip(String zipFileName) {
		ExcelBatchWriteContext context = createContext();
		context.setZipFileName(zipFileName);
		ExcelBatchWritePipelines.zipPipeline().execute(context);
		return this;
	}

	public ExcelLargeListBatchWriter writer() {
		if (writer == null) {
			if (maxPoolSize == null) {
				writer = ExcelWriterFactory.createLargeListBatchWriter(outputDirPath, rowAccessWindowSize,
						compressTempFiles, useSharedStringsTable);
			} else {
				writer = ExcelWriterFactory.createLargeListBatchWriter(outputDirPath, maxPoolSize,
						rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
			}
		}
		return writer;
	}

	private <T extends ExcelBaseModel> void addListOperation(List<T> modelList, Class<? extends ExcelBaseModel> modelClass, String[] excludeFields) {
		ExcelBatchWriteContext.ListOperation operation = new ExcelBatchWriteContext.ListOperation();
		operation.setModelList(modelList);
		operation.setFileName(fileName);
		operation.setSheetName(sheetName);
		operation.setModelClass(modelClass);
		operation.setExcludeFields(excludeFields);
		operations.add(operation);
	}

	private ExcelBatchWriteContext createContext() {
		ExcelBatchWriteContext context = new ExcelBatchWriteContext();
		context.setWriter(writer());
		context.setStyles(styles);
		context.setOperations(operations);
		return context;
	}
}
