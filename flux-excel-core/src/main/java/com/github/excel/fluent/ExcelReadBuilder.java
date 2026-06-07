package com.github.excel.fluent;

import com.github.excel.boot.ExcelMetadataRegistry;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReadError;
import com.github.excel.model.ExcelReadResult;
import com.github.excel.model.ExcelSheetInfo;
import com.github.excel.param.ExcelReaderFileParam;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.param.ExcelReaderModelParam;
import com.github.excel.param.ExcelReaderParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.read.facade.ExcelCustomReader;
import com.github.excel.read.facade.ExcelReaderBatchProcess;
import com.github.excel.read.format.ExcelDefaultReaderDataFormat;
import com.github.excel.read.format.ExcelReaderDataFormat;
import com.github.excel.read.handler.reader.ExcelReader;
import com.github.excel.read.handler.reader.ExcelReaderFactory;
import com.github.excel.read.handler.row.ExcelReaderRowHandler;
import com.github.excel.read.listener.ExcelReadListener;
import com.github.excel.read.listener.ExcelReadListenerContext;
import com.github.excel.util.DynamicExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelReadBuilder {

	private final ExcelReaderParam readerParam;
	private final ExcelReader reader;
	private final List<Class<? extends ExcelBaseModel>> csvListClasses = new ArrayList<>();
	private final Map<Class<? extends ExcelBaseModel>, List<? extends ExcelBaseModel>> csvListResultMap = new HashMap<>();
	private final ExcelReaderDataFormat csvDataFormat = new ExcelDefaultReaderDataFormat();
	private final Map<String, String> headerAliases = new LinkedHashMap<>();
	private final Set<Class<? extends ExcelBaseModel>> readTargetClasses = new LinkedHashSet<>();
	private List<Map<String, Object>> cachedMapList;
	private int sheetIndex = 0;
	private String sheetName;
	private Integer headRowNumber;
	private Integer dataStartRow;
	private Integer dataEndRow;
	private Integer maxRows;
	private ExcelReadListener<? extends ExcelBaseModel> readListener;
	private final Map<Class<? extends ExcelBaseModel>, Integer> listenerRowCounts = new HashMap<>();

	private ExcelReadBuilder(ExcelReaderParam readerParam) {
		this.readerParam = readerParam;
		if (readerParam instanceof ExcelReaderFileParam) {
			this.reader = ExcelReaderFactory.createUserReader((ExcelReaderFileParam) readerParam);
		} else {
			this.reader = ExcelReaderFactory.createUserReader((ExcelReaderStreamParam) readerParam);
		}
	}

	public static ExcelReadBuilder fromStream(InputStream inputStream) {
		return new ExcelReadBuilder(ExcelReaderStreamParam.builder()
				.stream(inputStream)
				.build());
	}

	public static ExcelReadBuilder fromFile(File file) {
		return new ExcelReadBuilder(ExcelReaderFileParam.builder()
				.file(file)
				.build());
	}

	public ExcelReadBuilder fileName(String fileName) {
		readerParam.setTemplate(fileName);
		return this;
	}

	public ExcelReadBuilder template(String template) {
		readerParam.setTemplate(template);
		return this;
	}

	public ExcelReadBuilder password(String password) {
		readerParam.setPassword(password);
		return this;
	}

	public ExcelReadBuilder closeInputStream(boolean closeInputStream) {
		readerParam.setCloseInputStream(closeInputStream);
		return this;
	}

	public ExcelReadBuilder readPicture(boolean readPicture) {
		readerParam.setReadPicture(readPicture);
		return this;
	}

	public ExcelReadBuilder readOnly(boolean readOnly) {
		readerParam.setReadOnly(readOnly);
		return this;
	}

	public ExcelReadBuilder failFast(boolean failFast) {
		reader.setFailFast(failFast);
		return this;
	}

	public ExcelReadBuilder sheet(int sheetIndex) {
		this.sheetIndex = sheetIndex;
		this.sheetName = null;
		return this;
	}

	public ExcelReadBuilder sheet(String sheetName) {
		this.sheetName = sheetName;
		return this;
	}

	public ExcelReadBuilder headRowNumber(int headRowNumber) {
		this.headRowNumber = headRowNumber;
		return this;
	}

	public ExcelReadBuilder dataStartRow(int dataStartRow) {
		this.dataStartRow = dataStartRow;
		return this;
	}

	public ExcelReadBuilder dataEndRow(int dataEndRow) {
		this.dataEndRow = dataEndRow;
		return this;
	}

	public ExcelReadBuilder maxRows(int maxRows) {
		this.maxRows = maxRows;
		return this;
	}

	public ExcelReadBuilder ignoreEmptyRow(boolean ignoreEmptyRow) {
		readerParam.setIgnoreEmptyRow(ignoreEmptyRow);
		return this;
	}

	public ExcelReadBuilder trimString(boolean trimString) {
		readerParam.setTrimString(trimString);
		return this;
	}

	public ExcelReadBuilder mergedCellStrategy(String mergedCellStrategy) {
		readerParam.setMergedCellStrategy(mergedCellStrategy);
		return this;
	}

	public ExcelReadBuilder emptyCellPolicy(String emptyCellPolicy) {
		readerParam.setEmptyCellPolicy(emptyCellPolicy);
		return this;
	}

	public ExcelReadBuilder headerAlias(String titleName, String fieldName) {
		headerAliases.put(titleName, fieldName);
		return this;
	}

	public ExcelReadBuilder converter(Class<?> type, java.util.function.Function<Object, Object> converter) {
		readerParam.getTypeConverters().put(type, converter);
		return this;
	}

	public ExcelReadBuilder fieldConverter(String field, java.util.function.Function<Object, Object> converter) {
		readerParam.getFieldConverters().put(field, converter);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder readListener(ExcelReadListener<T> readListener) {
		this.readListener = readListener;
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder model(Class<T> modelClass) {
		readTargetClasses.add(modelClass);
		reader.addModel(applyReadOptions(ExcelReaderModelParam.<T>builder()
				.modelCla(modelClass)
				.build()));
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder model(Class<T> modelClass, ExcelReaderRowHandler<T> rowHandler) {
		readTargetClasses.add(modelClass);
		reader.addModel(applyReadOptions(ExcelReaderModelParam.<T>builder()
				.modelCla(modelClass)
				.rowHandler(rowHandler)
				.build()));
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder list(Class<T> modelClass) {
		readTargetClasses.add(modelClass);
		csvListClasses.add(modelClass);
		reader.addList(applyReadOptions(ExcelReaderListParam.<T>builder()
				.modelCla(modelClass)
				.build()));
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder list(Class<T> modelClass, ExcelReaderRowHandler<T> rowHandler) {
		readTargetClasses.add(modelClass);
		csvListClasses.add(modelClass);
		reader.addList(applyReadOptions(ExcelReaderListParam.<T>builder()
				.modelCla(modelClass)
				.rowHandler(rowHandler)
				.build()));
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadBuilder list(Class<T> modelClass, ExcelReaderBatchProcess<T> batchProcess) {
		readTargetClasses.add(modelClass);
		csvListClasses.add(modelClass);
		reader.addList(applyReadOptions(ExcelReaderListParam.<T>builder()
				.modelCla(modelClass)
				.batchProcess(batchProcess)
				.build()));
		return this;
	}

	public ExcelReadBuilder custom(ExcelCustomReader customReader) {
		reader.readCustom(customReader);
		return this;
	}

	public ExcelReadBuilder parse() {
		if (isCsvFile() && !csvListClasses.isEmpty()) {
			parseCsvList();
			notifyRowsFromResults();
			notifyAfterAll();
			return this;
		}
		reader.parse();
		notifyRowsFromResults();
		notifyAfterAll();
		return this;
	}

	public <T extends ExcelBaseModel> ExcelReadResult<T> parseResult(Class<T> modelClass) {
		ExcelReadResult<T> result = new ExcelReadResult<>();
		readerParam.setCollectErrors(true);
		readerParam.getReadErrors().clear();
		try {
			parse();
			List<T> data = getList(modelClass);
			if ((data == null || data.isEmpty())) {
				T model = getModel(modelClass);
				if (model != null) {
					data = new ArrayList<>();
					data.add(model);
				}
			}
			if (data != null) {
				result.setData(data);
				result.setTotalRows(data.size());
				result.setValidRows(data.size());
			}
			if (readerParam.getReadErrors() != null) {
				result.getErrors().addAll(readerParam.getReadErrors());
				result.setInvalidRows(readerParam.getReadErrors().size());
			}
		} catch (Exception e) {
			ExcelReadError error = new ExcelReadError()
					.setSheetIndex(sheetIndex)
					.setSheetName(sheetName)
					.setMessage(e.getMessage());
			result.getErrors().add(error);
			result.setInvalidRows(1);
			notifyError(error);
		}
		return result;
	}

	public <T extends ExcelBaseModel> T getModel(Class<T> modelClass) {
		return reader.getModel(modelClass);
	}

	public <T extends ExcelBaseModel> List<T> getList(Class<T> modelClass) {
		if (csvListResultMap.containsKey(modelClass)) {
			return limitRows((List<T>) csvListResultMap.get(modelClass));
		}
		return limitRows(reader.getList(modelClass));
	}

	public List<Map<String, Object>> mapList() {
		if (cachedMapList != null) {
			return applyHeaderAliases(cachedMapList);
		}
		if (readerParam instanceof ExcelReaderFileParam) {
			try (InputStream inputStream = new FileInputStream(((ExcelReaderFileParam) readerParam).getFile())) {
				cachedMapList = DynamicExcelUtil.readMapList(inputStream, readerParam.getTemplate());
			} catch (FileNotFoundException e) {
				throw new IllegalStateException("Read excel map list file failed", e);
			} catch (IOException e) {
				throw new IllegalStateException("Read excel map list file failed", e);
			}
		} else {
			cachedMapList = DynamicExcelUtil.readMapList(((ExcelReaderStreamParam) readerParam).getStream(), readerParam.getTemplate());
		}
		return applyHeaderAliases(cachedMapList);
	}

	public List<ExcelSheetInfo> sheets() {
		if (!(readerParam instanceof ExcelReaderFileParam)) {
			return new ArrayList<>();
		}
		try (InputStream inputStream = new FileInputStream(((ExcelReaderFileParam) readerParam).getFile());
			 Workbook workbook = WorkbookFactory.create(inputStream)) {
			List<ExcelSheetInfo> sheetInfos = new ArrayList<>();
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				Sheet sheet = workbook.getSheetAt(i);
				sheetInfos.add(new ExcelSheetInfo(i, sheet.getSheetName(), sheet.getFirstRowNum(), sheet.getLastRowNum(), sheet.getPhysicalNumberOfRows()));
			}
			return sheetInfos;
		} catch (Exception e) {
			throw new IllegalStateException("Inspect excel sheets failed", e);
		}
	}

	public List<Map<String, Object>> preview(int rows) {
		List<Map<String, Object>> mapRows = mapList();
		return mapRows.size() <= rows ? mapRows : new ArrayList<>(mapRows.subList(0, rows));
	}

	public String getErrorMsg() {
		return reader.getErrorMsg();
	}

	public ExcelReader reader() {
		return reader;
	}

	private boolean isCsvFile() {
		return readerParam.getTemplate() != null && readerParam.getTemplate().toLowerCase().endsWith(".csv");
	}

	private void parseCsvList() {
		List<Map<String, Object>> rows = mapList();
		for (Class<? extends ExcelBaseModel> modelClass : csvListClasses) {
			csvListResultMap.put(modelClass, convertCsvRows(rows, modelClass));
		}
	}

	private <T extends ExcelBaseModel, P extends com.github.excel.param.ExcelReaderDataParam<T>> P applyReadOptions(P param) {
		param.setSheetIndex(sheetIndex);
		param.setSheetName(sheetName);
		param.setHeadRowNumber(headRowNumber);
		param.setDataStartRow(dataStartRow);
		param.setDataEndRow(dataEndRow);
		param.setMaxRows(maxRows);
		return param;
	}

	private <T extends ExcelBaseModel> void notifyRow(Class<T> modelClass, T model) {
		if (readListener == null) {
			return;
		}
		int notifiedRows = listenerRowCounts.getOrDefault(modelClass, 0);
		if (maxRows != null && notifiedRows >= maxRows) {
			return;
		}
		listenerRowCounts.put(modelClass, notifiedRows + 1);
		((ExcelReadListener<T>) readListener).onRow(model, new ExcelReadListenerContext()
				.setSheetIndex(sheetIndex)
				.setSheetName(sheetName));
	}

	private void notifyRowsFromResults() {
		if (readListener == null) {
			return;
		}
		for (Class<? extends ExcelBaseModel> targetClass : readTargetClasses) {
			notifyRows(targetClass);
		}
	}

	private <T extends ExcelBaseModel> void notifyRows(Class<T> targetClass) {
		List<T> rows = getList(targetClass);
		if (rows != null && !rows.isEmpty()) {
			for (T row : rows) {
				notifyRow(targetClass, row);
			}
			return;
		}
		T model = getModel(targetClass);
		if (model != null) {
			notifyRow(targetClass, model);
		}
	}

	private <T extends ExcelBaseModel> List<T> limitRows(List<T> rows) {
		if (rows == null || maxRows == null || rows.size() <= maxRows) {
			return rows;
		}
		return new ArrayList<>(rows.subList(0, maxRows));
	}

	private void notifyError(ExcelReadError error) {
		if (readListener == null) {
			return;
		}
		((ExcelReadListener) readListener).onError(error, new ExcelReadListenerContext()
				.setSheetIndex(sheetIndex)
				.setSheetName(sheetName));
	}

	private void notifyAfterAll() {
		if (readListener == null) {
			return;
		}
		((ExcelReadListener) readListener).afterAll(new ExcelReadListenerContext()
				.setSheetIndex(sheetIndex)
				.setSheetName(sheetName));
	}

	private List<Map<String, Object>> applyHeaderAliases(List<Map<String, Object>> rows) {
		if (headerAliases.isEmpty() || rows == null) {
			return rows;
		}
		List<Map<String, Object>> result = new ArrayList<>();
		for (Map<String, Object> row : rows) {
			Map<String, Object> mapped = new LinkedHashMap<>();
			for (Map.Entry<String, Object> entry : row.entrySet()) {
				mapped.put(headerAliases.getOrDefault(entry.getKey(), entry.getKey()), entry.getValue());
			}
			result.add(mapped);
		}
		cachedMapList = result;
		return result;
	}

	private <T extends ExcelBaseModel> List<T> convertCsvRows(List<Map<String, Object>> rows, Class<T> modelClass) {
		ExcelCacheImportModel cacheModel = ExcelMetadataRegistry.getExcelCacheImportMapValue(modelClass);
		List<T> result = new ArrayList<>();
		for (Map<String, Object> row : rows) {
			try {
				T model = modelClass.getDeclaredConstructor().newInstance();
				for (Map.Entry<String, ExcelCacheImportModel.ExcelCacheImportFieldModel> entry : cacheModel.getFieldModelMap().entrySet()) {
					if (!row.containsKey(entry.getKey())) {
						continue;
					}
					ExcelCacheImportModel.ExcelCacheImportFieldModel fieldModel = entry.getValue();
					Class<?> targetType = fieldModel.getSetMethod().getParameterTypes()[0];
					Object value = csvDataFormat.format(row.get(entry.getKey()), fieldModel.getImportProperty().formatPattern(), wrapPrimitive(targetType));
					fieldModel.getSetMethod().invoke(model, value);
				}
				result.add(model);
			} catch (InstantiationException | IllegalAccessException | InvocationTargetException | NoSuchMethodException | ParseException e) {
				throw new IllegalStateException("Read csv model list failed", e);
			}
		}
		return result;
	}

	private Class<?> wrapPrimitive(Class<?> targetType) {
		if (!targetType.isPrimitive()) {
			return targetType;
		}
		if (targetType == int.class) {
			return Integer.class;
		}
		if (targetType == long.class) {
			return Long.class;
		}
		if (targetType == short.class) {
			return Short.class;
		}
		if (targetType == byte.class) {
			return Byte.class;
		}
		if (targetType == float.class) {
			return Float.class;
		}
		if (targetType == double.class) {
			return Double.class;
		}
		if (targetType == boolean.class) {
			return Boolean.class;
		}
		return targetType;
	}
}
