package com.github.excel.fluent;

import com.github.excel.boot.ExcelMetadataRegistry;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCellComment;
import com.github.excel.model.ExcelCellHyperlink;
import com.github.excel.model.ExcelHeaderInfo;
import com.github.excel.model.ExcelMergedCell;
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
import com.github.excel.read.pipeline.execution.ExcelReadExecutionContext;
import com.github.excel.read.pipeline.execution.ExcelReadExecutionPipelines;
import com.github.excel.util.DynamicExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

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
import java.util.concurrent.atomic.AtomicInteger;

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
		ExcelReadExecutionContext context = ExcelReadExecutionContext.builder()
				.csvFile(isCsvFile())
				.csvListRead(!csvListClasses.isEmpty())
				.csvParser(this::parseCsvList)
				.workbookParser(reader::parse)
				.rowNotifier(this::notifyRowsFromResults)
				.afterAllNotifier(this::notifyAfterAll)
				.build();
		ExcelReadExecutionPipelines.fluentReadPipeline().execute(context);
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
				cachedMapList = DynamicExcelUtil.readMapList(inputStream, readerParam.getTemplate(), sheetIndex, sheetName,
						headRowNumber, dataStartRow, dataEndRow, maxRows,
						Boolean.TRUE.equals(readerParam.getTrimString()),
						Boolean.TRUE.equals(readerParam.getIgnoreEmptyRow()), true);
			} catch (FileNotFoundException e) {
				throw new IllegalStateException("Read excel map list file failed", e);
			} catch (IOException e) {
				throw new IllegalStateException("Read excel map list file failed", e);
			}
		} else {
			cachedMapList = DynamicExcelUtil.readMapList(((ExcelReaderStreamParam) readerParam).getStream(), readerParam.getTemplate(),
					sheetIndex, sheetName, headRowNumber, dataStartRow, dataEndRow, maxRows,
					Boolean.TRUE.equals(readerParam.getTrimString()),
					Boolean.TRUE.equals(readerParam.getIgnoreEmptyRow()), true);
		}
		return applyHeaderAliases(cachedMapList);
	}

	public List<String> headers() {
		List<String> headers = new ArrayList<>();
		for (ExcelHeaderInfo headerInfo : readHeaders()) {
			headers.add(headerInfo.getTitle());
		}
		return headers;
	}

	public List<ExcelHeaderInfo> readHeaders() {
		return inspectWorkbook(workbook -> {
			Sheet sheet = resolveSheet(workbook);
			if (sheet == null) {
				return new ArrayList<>();
			}
			DataFormatter formatter = new DataFormatter();
			int headerRowIndex = headRowNumber == null ? sheet.getFirstRowNum() : headRowNumber;
			Row headerRow = sheet.getRow(headerRowIndex);
			List<ExcelHeaderInfo> headers = new ArrayList<>();
			for (int colIndex = 0; headerRow != null && colIndex < headerRow.getLastCellNum(); colIndex++) {
				Cell cell = headerRow.getCell(colIndex);
				String title = formatter.formatCellValue(cell);
				if (Boolean.TRUE.equals(readerParam.getTrimString()) && title != null) {
					title = title.trim();
				}
				headers.add(new ExcelHeaderInfo()
						.setSheetIndex(workbook.getSheetIndex(sheet))
						.setSheetName(sheet.getSheetName())
						.setRowIndex(headerRowIndex)
						.setColIndex(colIndex)
						.setTitle(title));
			}
			return headers;
		});
	}

	public List<ExcelCellComment> comments() {
		return readComments();
	}

	public List<ExcelCellComment> readComments() {
		return inspectWorkbook(workbook -> {
			DataFormatter formatter = new DataFormatter();
			List<ExcelCellComment> comments = new ArrayList<>();
			for (Sheet sheet : resolveSheets(workbook)) {
				for (Row row : sheet) {
					for (Cell cell : row) {
						Comment comment = cell.getCellComment();
						if (comment == null) {
							continue;
						}
						comments.add(new ExcelCellComment()
								.setSheetIndex(workbook.getSheetIndex(sheet))
								.setSheetName(sheet.getSheetName())
								.setRowIndex(cell.getRowIndex())
								.setColIndex(cell.getColumnIndex())
								.setCellRef(cell.getAddress().formatAsString())
								.setAuthor(comment.getAuthor())
								.setText(comment.getString() == null ? null : comment.getString().getString())
								.setCellValue(formatter.formatCellValue(cell)));
					}
				}
			}
			return comments;
		});
	}

	public List<ExcelCellHyperlink> hyperlinks() {
		return readHyperlinks();
	}

	public List<ExcelCellHyperlink> readHyperlinks() {
		return inspectWorkbook(workbook -> {
			DataFormatter formatter = new DataFormatter();
			List<ExcelCellHyperlink> hyperlinks = new ArrayList<>();
			for (Sheet sheet : resolveSheets(workbook)) {
				for (Row row : sheet) {
					for (Cell cell : row) {
						Hyperlink hyperlink = cell.getHyperlink();
						if (hyperlink == null) {
							continue;
						}
						hyperlinks.add(new ExcelCellHyperlink()
								.setSheetIndex(workbook.getSheetIndex(sheet))
								.setSheetName(sheet.getSheetName())
								.setRowIndex(cell.getRowIndex())
								.setColIndex(cell.getColumnIndex())
								.setCellRef(cell.getAddress().formatAsString())
								.setAddress(hyperlink.getAddress())
								.setLabel(hyperlink.getLabel())
								.setType(hyperlink.getType() == null ? null : hyperlink.getType().name())
								.setCellValue(formatter.formatCellValue(cell)));
					}
				}
			}
			return hyperlinks;
		});
	}

	public List<ExcelMergedCell> mergedCells() {
		return readMergedCells();
	}

	public List<ExcelMergedCell> readMergedCells() {
		return inspectWorkbook(workbook -> {
			DataFormatter formatter = new DataFormatter();
			List<ExcelMergedCell> mergedCells = new ArrayList<>();
			for (Sheet sheet : resolveSheets(workbook)) {
				for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
					CellRangeAddress range = sheet.getMergedRegion(i);
					Row firstRow = sheet.getRow(range.getFirstRow());
					Cell firstCell = firstRow == null ? null : firstRow.getCell(range.getFirstColumn());
					mergedCells.add(new ExcelMergedCell()
							.setSheetIndex(workbook.getSheetIndex(sheet))
							.setSheetName(sheet.getSheetName())
							.setFirstRow(range.getFirstRow())
							.setLastRow(range.getLastRow())
							.setFirstCol(range.getFirstColumn())
							.setLastCol(range.getLastColumn())
							.setCellRange(range.formatAsString())
							.setValue(formatter.formatCellValue(firstCell)));
				}
			}
			return mergedCells;
		});
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

	private interface WorkbookInspector<R> {
		R inspect(Workbook workbook) throws Exception;
	}

	private <R> R inspectWorkbook(WorkbookInspector<R> inspector) {
		if (isCsvFile()) {
			throw new IllegalStateException("CSV files do not contain workbook metadata");
		}
		try (InputStream inputStream = openInputStream();
			 Workbook workbook = WorkbookFactory.create(inputStream)) {
			return inspector.inspect(workbook);
		} catch (Exception e) {
			throw new IllegalStateException("Inspect excel workbook failed", e);
		}
	}

	private InputStream openInputStream() throws FileNotFoundException {
		if (readerParam instanceof ExcelReaderFileParam) {
			return new FileInputStream(((ExcelReaderFileParam) readerParam).getFile());
		}
		return ((ExcelReaderStreamParam) readerParam).getStream();
	}

	private Sheet resolveSheet(Workbook workbook) {
		if (sheetName != null && !sheetName.trim().isEmpty()) {
			return workbook.getSheet(sheetName);
		}
		return workbook.getSheetAt(sheetIndex);
	}

	private List<Sheet> resolveSheets(Workbook workbook) {
		List<Sheet> sheets = new ArrayList<>();
		Sheet selectedSheet = resolveSheet(workbook);
		if (selectedSheet != null) {
			sheets.add(selectedSheet);
		}
		return sheets;
	}

	private void parseCsvList() {
		for (Class<? extends ExcelBaseModel> modelClass : csvListClasses) {
			csvListResultMap.put(modelClass, new ArrayList<>());
		}
		try (InputStream inputStream = openInputStream()) {
			AtomicInteger dataRowIndex = new AtomicInteger(0);
			AtomicInteger acceptedRows = new AtomicInteger(0);
			DynamicExcelUtil.readCsv(inputStream, row -> {
				int currentDataRow = dataRowIndex.incrementAndGet();
				if (!csvRowInRange(currentDataRow, acceptedRows.get())) {
					return;
				}
				Map<String, Object> mappedRow = applyHeaderAliases(row);
				for (Class<? extends ExcelBaseModel> modelClass : csvListClasses) {
					List<ExcelBaseModel> modelRows = (List<ExcelBaseModel>) csvListResultMap.get(modelClass);
					modelRows.add(convertCsvRow(mappedRow, modelClass));
				}
				acceptedRows.incrementAndGet();
			});
		} catch (FileNotFoundException e) {
			throw new IllegalStateException("Read csv model list file failed", e);
		} catch (IOException e) {
			throw new IllegalStateException("Read csv model list file failed", e);
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

	private Map<String, Object> applyHeaderAliases(Map<String, Object> row) {
		if (headerAliases.isEmpty() || row == null) {
			return row;
		}
		Map<String, Object> mapped = new LinkedHashMap<>();
		for (Map.Entry<String, Object> entry : row.entrySet()) {
			mapped.put(headerAliases.getOrDefault(entry.getKey(), entry.getKey()), entry.getValue());
		}
		return mapped;
	}

	private <T extends ExcelBaseModel> List<T> convertCsvRows(List<Map<String, Object>> rows, Class<T> modelClass) {
		ExcelCacheImportModel cacheModel = ExcelMetadataRegistry.getExcelCacheImportMapValue(modelClass);
		List<T> result = new ArrayList<>();
		for (Map<String, Object> row : rows) {
			result.add(convertCsvRow(row, modelClass, cacheModel));
		}
		return result;
	}

	private <T extends ExcelBaseModel> T convertCsvRow(Map<String, Object> row, Class<T> modelClass) {
		return convertCsvRow(row, modelClass, ExcelMetadataRegistry.getExcelCacheImportMapValue(modelClass));
	}

	private <T extends ExcelBaseModel> T convertCsvRow(Map<String, Object> row, Class<T> modelClass,
													   ExcelCacheImportModel cacheModel) {
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
			return model;
		} catch (InstantiationException | IllegalAccessException | InvocationTargetException | NoSuchMethodException | ParseException e) {
			throw new IllegalStateException("Read csv model list failed", e);
		}
	}

	private boolean csvRowInRange(int dataRowIndex, int acceptedRows) {
		if (dataStartRow != null && dataRowIndex < dataStartRow) {
			return false;
		}
		if (dataEndRow != null && dataRowIndex > dataEndRow) {
			return false;
		}
		return maxRows == null || acceptedRows < maxRows;
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
