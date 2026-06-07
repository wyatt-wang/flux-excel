package com.github.excel.fluent;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.enums.ExcelWriterFillStyleEnum;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.boot.ExcelMetadataRegistry;
import com.github.excel.model.ExcelCacheFieldModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.param.ExcelWriterCellParam;
import com.github.excel.param.ExcelWriterComboParam;
import com.github.excel.param.ExcelWriterCommentParam;
import com.github.excel.param.ExcelWriterConditionalStyleParam;
import com.github.excel.param.ExcelWriterFileParam;
import com.github.excel.param.ExcelWriterListParam;
import com.github.excel.param.ExcelWriterMergeParam;
import com.github.excel.param.ExcelWriterModelParam;
import com.github.excel.param.ExcelWriterNumberScopeParam;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.param.ExcelWriterSteamParam;
import com.github.excel.write.ExcelCustomWriter;
import com.github.excel.write.ExcelCellStyleConfigurer;
import com.github.excel.write.ExcelTemplateFillParam;
import com.github.excel.write.ExcelWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.util.DynamicExcelUtil;
import com.github.excel.util.ExcelWatermarkUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWriteBuilder {

	private enum TargetType {
		STREAM, FILE
	}

	private final TargetType targetType;
	private final OutputStream outputStream;
	private final File file;
	private final ExcelWriter writer;
	private final ExcelWriterParam writerParam;
	private String sheetName = "sheet";
	private String fileName = "export";
	private ExcelSuffixEnum suffix = ExcelSuffixEnum.XLSX;
	private int rowIndex = 0;
	private int colIndex = 0;
	private boolean waterfall = false;
	private boolean fillTemplate = true;
	private List<? extends Map<String, ?>> mapList;
	private Map<String, String> emptyHeaders;
	private final List<Map<String, Object>> csvModelRows = new ArrayList<>();
	private Map<String, String> csvModelHeaders;
	private final Map<String, String> headerAliases = new LinkedHashMap<>();
	private final Map<Class<?>, Function<Object, Object>> typeConverters = new LinkedHashMap<>();
	private final Map<String, Function<Object, Object>> fieldConverters = new LinkedHashMap<>();
	private final List<ExcelTemplateFillParam> templateFillParams = new ArrayList<>();
	private final Set<String> includeFields = new LinkedHashSet<>();
	private boolean onlyAlias = false;
	private boolean dynamicTable = false;
	private Class<? extends ExcelBaseModel> tableModelClass;
	private ExcelCellStyleConfigurer headerStyleConfigurer;
	private final Map<String, ExcelCellStyleConfigurer> columnStyleConfigurers = new LinkedHashMap<>();
	private String watermarkText;
	private ExcelCustomWriter customWriter;

	private ExcelWriteBuilder(TargetType targetType, OutputStream outputStream, File file) {
		this.targetType = targetType;
		this.outputStream = outputStream;
		this.file = file;
		if (file != null) {
			this.fileName = file.getName();
			this.suffix = resolveSuffix(file.getName());
		}
		this.writerParam = createWriterParam();
		this.writer = ExcelWriterFactory.createUserModelWriter(writerParam);
	}

	public static ExcelWriteBuilder toStream(OutputStream outputStream) {
		return new ExcelWriteBuilder(TargetType.STREAM, outputStream, null);
	}

	public static ExcelWriteBuilder toFile(File file) {
		return new ExcelWriteBuilder(TargetType.FILE, null, file);
	}

	public ExcelWriteBuilder fileName(String fileName) {
		this.fileName = fileName;
		return this;
	}

	public ExcelWriteBuilder suffix(ExcelSuffixEnum suffix) {
		this.suffix = suffix;
		writerParam.setSuffixEnum(suffix);
		return this;
	}

	public ExcelWriteBuilder template(String template) {
		writerParam.setTemplate(template);
		return this;
	}

	public ExcelWriteBuilder templateFilePath(String templateFilePath) {
		writerParam.setTemplateFilePath(templateFilePath);
		return this;
	}

	public ExcelWriteBuilder templatePassword(String templatePassword) {
		writerParam.setTemplatePassword(templatePassword);
		return this;
	}

	public ExcelWriteBuilder streaming(boolean streaming) {
		writer.setStreaming(streaming);
		return this;
	}

	public ExcelWriteBuilder rowAccessWindowSize(int rowAccessWindowSize) {
		writerParam.setRowAccessWindowSize(rowAccessWindowSize);
		return this;
	}

	public ExcelWriteBuilder compressTempFiles(boolean compressTempFiles) {
		writerParam.setCompressTempFiles(compressTempFiles);
		return this;
	}

	public ExcelWriteBuilder useSharedStringsTable(boolean useSharedStringsTable) {
		writerParam.setUseSharedStringsTable(useSharedStringsTable);
		return this;
	}

	public ExcelWriteBuilder noneDataTips(boolean noneDataTips) {
		writer.setNoneDataTips(noneDataTips);
		return this;
	}

	public ExcelWriteBuilder noneDataTipsMsg(String noneDataTipsMsg) {
		writerParam.setNoneDataTipsMsg(noneDataTipsMsg);
		return this;
	}

	public ExcelWriteBuilder sheetRowMaxCount(int sheetRowMaxCount) {
		writerParam.setSheetRowMaxCount(sheetRowMaxCount);
		return this;
	}

	public ExcelWriteBuilder sheet(String sheetName) {
		this.sheetName = sheetName;
		writer.selectSheet(sheetName);
		return this;
	}

	public ExcelWriteBuilder table(Class<? extends ExcelBaseModel> modelClass) {
		this.tableModelClass = modelClass;
		return this;
	}

	public ExcelWriteBuilder at(int rowIndex, int colIndex) {
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		return this;
	}

	public ExcelWriteBuilder waterfall() {
		return waterfall(true);
	}

	public ExcelWriteBuilder waterfall(boolean waterfall) {
		this.waterfall = waterfall;
		return this;
	}

	public ExcelWriteBuilder fillTemplate(boolean fillTemplate) {
		this.fillTemplate = fillTemplate;
		return this;
	}

	public ExcelWriteBuilder alias(String field, String titleName) {
		headerAliases.put(field, titleName);
		dynamicTable = true;
		return this;
	}

	public ExcelWriteBuilder onlyAlias(boolean onlyAlias) {
		this.onlyAlias = onlyAlias;
		dynamicTable = true;
		return this;
	}

	public ExcelWriteBuilder include(String... fields) {
		if (fields != null) {
			includeFields.addAll(Arrays.asList(fields));
		}
		dynamicTable = true;
		return this;
	}

	public ExcelWriteBuilder converter(Class<?> type, Function<Object, Object> converter) {
		typeConverters.put(type, converter);
		dynamicTable = true;
		return this;
	}

	public ExcelWriteBuilder fieldConverter(String field, Function<Object, Object> converter) {
		fieldConverters.put(field, converter);
		dynamicTable = true;
		return this;
	}

	public ExcelWriteBuilder headerStyle(ExcelCellStyleConfigurer styleConfigurer) {
		this.headerStyleConfigurer = styleConfigurer;
		return this;
	}

	public ExcelWriteBuilder columnStyle(String field, ExcelCellStyleConfigurer styleConfigurer) {
		columnStyleConfigurers.put(field, styleConfigurer);
		return this;
	}

	public ExcelWriteBuilder fill(String name, Object value) {
		templateFillParams.add(new ExcelTemplateFillParam(name, value, null));
		return this;
	}

	public ExcelWriteBuilder fillList(String name, List<?> values) {
		templateFillParams.add(new ExcelTemplateFillParam(name, null, values));
		return this;
	}

	public <T extends ExcelBaseModel> ExcelWriteBuilder list(List<T> modelList) {
		if (tableModelClass == null) {
			throw new ExcelWriterException("table model class required before list");
		}
		return list(modelList, (Class<T>) tableModelClass);
	}

	public <T extends ExcelBaseModel> ExcelWriteBuilder model(T model, Class<T> modelClass) {
		writer.writeModel(ExcelWriterModelParam.<T>builder()
				.model(model)
				.modelCla(modelClass)
				.sheetName(sheetName)
				.rowIndex(rowIndex)
				.colIndex(colIndex)
				.fillTemplate(fillTemplate)
				.build());
		appendCsvModelRows(Arrays.asList(model), modelClass);
		advancePositionAfterModel(modelClass);
		return this;
	}

	public <T extends ExcelBaseModel> ExcelWriteBuilder list(List<T> modelList, Class<T> modelClass) {
		writer.writeList(ExcelWriterListParam.<T>builder()
				.modelList(modelList)
				.modelCla(modelClass)
				.sheetName(sheetName)
				.rowIndex(rowIndex)
				.colIndex(colIndex)
				.fillTemplate(fillTemplate)
				.build());
		appendCsvModelRows(modelList, modelClass);
		advancePositionAfterList(modelList, modelClass);
		return this;
	}

	public ExcelWriteBuilder listMap(List<? extends Map<String, ?>> rows) {
		this.mapList = rows;
		return this;
	}

	public ExcelWriteBuilder emptyHeader(Map<String, String> headers) {
		this.emptyHeaders = headers;
		writer.setNoneDataTips(false);
		return this;
	}

	public ExcelWriteBuilder emptyHeader(Class<? extends ExcelBaseModel> modelClass) {
		ExcelCacheModel cacheModel = ExcelMetadataRegistry.getExcelCacheMapValue(modelClass);
		Map<String, String> headers = new LinkedHashMap<>();
		for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
			if (Objects.nonNull(fieldModel.getGetMethod())) {
				headers.put(fieldModel.getFieldName(), fieldModel.getTitleName());
			}
		}
		return emptyHeader(headers);
	}

	public ExcelWriteBuilder watermark(String text) {
		this.watermarkText = text;
		return this;
	}

	public ExcelWriteBuilder exclude(String... fields) {
		writer.excludes(fields);
		return this;
	}

	public ExcelWriteBuilder comment(String field, ExcelWriterCommentParam commentParam) {
		writer.addValidationOrComment(field, commentParam);
		return this;
	}

	public ExcelWriteBuilder validation(String field, ExcelWriterNumberScopeParam validationParam) {
		writer.addValidationOrComment(field, validationParam);
		return this;
	}

	public ExcelWriteBuilder validation(String field, ExcelWriterComboParam validationParam) {
		writer.addValidationOrComment(field, validationParam);
		return this;
	}

	public ExcelWriteBuilder conditionalStyle(String field, ExcelWriterConditionalStyleParam conditionalStyleParam) {
		writer.addConditionalStyle(field, conditionalStyleParam);
		return this;
	}

	public ExcelWriteBuilder conditionalStyle(ExcelWriterConditionalStyleParam conditionalStyleParam) {
		if (conditionalStyleParam != null && (conditionalStyleParam.getSheetName() == null || conditionalStyleParam.getSheetName().trim().isEmpty())) {
			conditionalStyleParam.setSheetName(sheetName);
		}
		writer.addConditionalStyle(conditionalStyleParam);
		return this;
	}

	public ExcelWriteBuilder column(ExcelWriterCellParam cellParam) {
		writer.writeColumn(cellParam);
		return this;
	}

	public ExcelWriteBuilder mergeColumn(ExcelWriterMergeParam mergeParam) {
		writer.writeMergeColumn(mergeParam);
		return this;
	}

	public ExcelWriteBuilder custom(ExcelCustomWriter customWriter) {
		this.customWriter = customWriter;
		return this;
	}

	@SafeVarargs
	public final ExcelWriteBuilder style(Class<? extends AbstractExcelStyle>... styles) {
		Arrays.stream(styles).forEach(writer::addStyle);
		return this;
	}

	public ExcelWriter writer() {
		return writer;
	}

	public ExcelWriteBuilder export() {
		if (Objects.nonNull(mapList) || Objects.nonNull(emptyHeaders) || (ExcelSuffixEnum.CSV == suffix && !csvModelRows.isEmpty())) {
			exportDynamic();
			return this;
		}
		applyRuntimeOptionsToWriter();
		applyCustomWriter();
		try (OutputStream stream = openTargetStream()) {
			writer.export(stream, fileName, suffix);
			return this;
		} catch (IOException e) {
			throw new ExcelWriterException(e);
		}
	}

	private void exportDynamic() {
		try (OutputStream targetStream = openTargetStream()) {
			List<? extends Map<String, ?>> rows = Objects.nonNull(mapList) ? mapList : csvModelRows;
			Map<String, String> headers = Objects.nonNull(emptyHeaders) ? emptyHeaders : csvModelHeaders;
			DynamicExcelUtil.writeMapList(targetStream, suffix, sheetName, rows, headers, watermarkText);
		} catch (IOException e) {
			throw new ExcelWriterException(e);
		}
	}

	private void applyRuntimeOptionsToWriter() {
		writer.runtimeHeaderAliases(headerAliases)
				.runtimeIncludeFields(includeFields)
				.runtimeOnlyAlias(onlyAlias)
				.runtimeTypeConverters(typeConverters)
				.runtimeFieldConverters(fieldConverters);
	}

	private OutputStream openTargetStream() throws IOException {
		if (targetType == TargetType.STREAM) {
			return outputStream;
		}
		return new FileOutputStream(file);
	}

	private void applyCustomWriter() {
		boolean hasPostProcessor = Objects.nonNull(customWriter)
				|| !templateFillParams.isEmpty()
				|| Objects.nonNull(headerStyleConfigurer)
				|| !columnStyleConfigurers.isEmpty()
				|| (watermarkText != null && !watermarkText.trim().isEmpty());
		if (!hasPostProcessor) {
			return;
		}
		writer.writeCustom(workbook -> {
			if (Objects.nonNull(customWriter)) {
				customWriter.execute(workbook);
			}
			applyTemplateFill(workbook);
			applyLightStyles(workbook);
			if (watermarkText != null && !watermarkText.trim().isEmpty()) {
				ExcelWatermarkUtil.applyTextWatermark(workbook, watermarkText);
			}
		});
	}

	private <T extends ExcelBaseModel> void appendCsvModelRows(List<T> rows, Class<T> modelClass) {
		ExcelCacheModel cacheModel = ExcelMetadataRegistry.getExcelCacheMapValue(modelClass);
		if (csvModelHeaders == null) {
			csvModelHeaders = new LinkedHashMap<>();
			for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
				String fieldName = fieldModel.getFieldName();
				if (skipRuntimeField(fieldName)) {
					continue;
				}
				csvModelHeaders.put(fieldName, headerAliases.getOrDefault(fieldName, fieldModel.getTitleName()));
			}
		}
		if (rows == null) {
			return;
		}
		long sequenceNo = 0L;
		for (T row : rows) {
			Map<String, Object> rowData = new LinkedHashMap<>();
			for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
				String fieldName = fieldModel.getFieldName();
				if (skipRuntimeField(fieldName)) {
					continue;
				}
				try {
					Object value = Objects.nonNull(fieldModel.getGetMethod())
							? fieldModel.getGetMethod().invoke(row)
							: ++sequenceNo;
					rowData.put(fieldName, convertRuntimeValue(fieldName, value));
				} catch (IllegalAccessException | InvocationTargetException e) {
					throw new ExcelWriterException(e);
				}
			}
			csvModelRows.add(rowData);
		}
	}

	private boolean skipRuntimeField(String fieldName) {
		if (!includeFields.isEmpty() && !includeFields.contains(fieldName)) {
			return true;
		}
		return onlyAlias && !headerAliases.containsKey(fieldName);
	}

	private Object convertRuntimeValue(String fieldName, Object value) {
		Function<Object, Object> fieldConverter = fieldConverters.get(fieldName);
		if (fieldConverter != null) {
			return fieldConverter.apply(value);
		}
		if (value == null) {
			return null;
		}
		for (Map.Entry<Class<?>, Function<Object, Object>> entry : typeConverters.entrySet()) {
			if (entry.getKey().isAssignableFrom(value.getClass())) {
				return entry.getValue().apply(value);
			}
		}
		return value;
	}

	private void applyTemplateFill(Workbook workbook) {
		if (templateFillParams.isEmpty()) {
			return;
		}
		Map<String, Object> values = new LinkedHashMap<>();
		Map<String, List<?>> listValues = new LinkedHashMap<>();
		for (ExcelTemplateFillParam fillParam : templateFillParams) {
			if (fillParam.getValue() != null) {
				values.put(fillParam.getName(), fillParam.getValue());
			}
			if (fillParam.getListValue() != null) {
				listValues.put(fillParam.getName(), fillParam.getListValue());
			}
		}
		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			for (Row row : sheet) {
				for (Cell cell : row) {
					if (cell.getCellType() != CellType.STRING) {
						continue;
					}
					String text = cell.getStringCellValue();
					boolean listFilled = false;
					for (Map.Entry<String, List<?>> entry : listValues.entrySet()) {
						String placeholder = "{" + entry.getKey() + "}";
						if (!text.contains(placeholder)) {
							continue;
						}
						List<?> list = entry.getValue();
						for (int index = 0; index < list.size(); index++) {
							Row targetRow = sheet.getRow(row.getRowNum() + index);
							if (targetRow == null) {
								targetRow = sheet.createRow(row.getRowNum() + index);
							}
							Cell targetCell = targetRow.getCell(cell.getColumnIndex());
							if (targetCell == null) {
								targetCell = targetRow.createCell(cell.getColumnIndex());
							}
							targetCell.setCellValue(text.replace(placeholder, Objects.toString(list.get(index), "")));
						}
						listFilled = true;
						break;
					}
					if (listFilled) {
						continue;
					}
					for (Map.Entry<String, Object> entry : values.entrySet()) {
						text = text.replace("{" + entry.getKey() + "}", Objects.toString(entry.getValue(), ""));
					}
					cell.setCellValue(text);
				}
			}
		}
	}

	private void applyLightStyles(Workbook workbook) {
		if (headerStyleConfigurer == null && columnStyleConfigurers.isEmpty()) {
			return;
		}
		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			Row headerRow = sheet.getRow(0);
			if (headerRow == null) {
				continue;
			}
			if (headerStyleConfigurer != null) {
				Map<Short, CellStyle> headerStyleCache = new LinkedHashMap<>();
				for (Cell cell : headerRow) {
					CellStyle sourceStyle = cell.getCellStyle();
					CellStyle cellStyle = headerStyleCache.computeIfAbsent(styleIndex(sourceStyle), key -> {
						CellStyle newStyle = workbook.createCellStyle();
						if (sourceStyle != null) {
							newStyle.cloneStyleFrom(sourceStyle);
						}
						headerStyleConfigurer.configure(newStyle, workbook);
						return newStyle;
					});
					cell.setCellStyle(cellStyle);
				}
			}
			for (Map.Entry<String, ExcelCellStyleConfigurer> entry : columnStyleConfigurers.entrySet()) {
				int columnIndex = resolveColumnIndex(headerRow, entry.getKey());
				if (columnIndex < 0) {
					continue;
				}
				Map<Short, CellStyle> columnStyleCache = new LinkedHashMap<>();
				for (int rowIndex = sheet.getFirstRowNum() + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					Row row = sheet.getRow(rowIndex);
					if (row == null) {
						continue;
					}
					Cell cell = row.getCell(columnIndex);
					if (cell == null) {
						cell = row.createCell(columnIndex);
					}
					CellStyle sourceStyle = cell.getCellStyle();
					CellStyle cellStyle = columnStyleCache.computeIfAbsent(styleIndex(sourceStyle), key -> {
						CellStyle newStyle = workbook.createCellStyle();
						if (sourceStyle != null) {
							newStyle.cloneStyleFrom(sourceStyle);
						}
						entry.getValue().configure(newStyle, workbook);
						return newStyle;
					});
					cell.setCellStyle(cellStyle);
				}
			}
		}
	}

	private short styleIndex(CellStyle cellStyle) {
		return cellStyle == null ? -1 : cellStyle.getIndex();
	}

	private int resolveColumnIndex(Row headerRow, String fieldOrTitleName) {
		String titleName = headerAliases.getOrDefault(fieldOrTitleName, fieldOrTitleName);
		for (Cell cell : headerRow) {
			if (cell.getCellType() == CellType.STRING && titleName.equals(cell.getStringCellValue())) {
				return cell.getColumnIndex();
			}
		}
		return -1;
	}

	private ExcelWriterParam createWriterParam() {
		if (targetType == TargetType.FILE) {
			return ExcelWriterFileParam.builder()
					.file(file)
					.suffixEnum(suffix)
					.build();
		}
		return ExcelWriterSteamParam.builder()
				.outputStream(outputStream)
				.suffixEnum(suffix)
				.build();
	}

	private void resetPosition() {
		rowIndex = 0;
		colIndex = 0;
	}

	private void advancePositionAfterModel(Class<? extends ExcelBaseModel> modelClass) {
		if (!waterfall) {
			resetPosition();
			return;
		}
		rowIndex += estimateModelRowSpan(modelClass);
	}

	private void advancePositionAfterList(List<? extends ExcelBaseModel> modelList, Class<? extends ExcelBaseModel> modelClass) {
		if (!waterfall) {
			resetPosition();
			return;
		}
		rowIndex += estimateListRowSpan(modelList, modelClass);
	}

	private int estimateListRowSpan(List<? extends ExcelBaseModel> modelList, Class<? extends ExcelBaseModel> modelClass) {
		ExcelWrite excelWrite = ExcelMetadataRegistry.getExcelCacheMapValue(modelClass).getExcelWrite();
		int rowCount = Objects.isNull(modelList) ? 0 : modelList.size();
		if (ExcelWriterFillStyleEnum.HORIZONTAL == excelWrite.fillStyle()) {
			return estimateHorizontalFieldRows(modelClass, excelWrite);
		}
		int titleRows = excelWrite.fillTitle() ? excelWrite.mergeTitleRowNum() + 1 : 0;
		int contentRows = rowCount * (excelWrite.mergeContentRowNum() + 1);
		return Math.max(1, titleRows + contentRows);
	}

	private int estimateModelRowSpan(Class<? extends ExcelBaseModel> modelClass) {
		ExcelCacheModel cacheModel = ExcelMetadataRegistry.getExcelCacheMapValue(modelClass);
		ExcelWrite excelWrite = cacheModel.getExcelWrite();
		if (ExcelWriterFillStyleEnum.HORIZONTAL == excelWrite.fillStyle()) {
			return estimateHorizontalFieldRows(modelClass, excelWrite);
		}
		int rows = 0;
		for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
			if (Objects.nonNull(fieldModel.getExportCell())) {
				rows += fieldModel.getExportCell().mergeRowNum() + 1;
			}
		}
		return Math.max(1, rows);
	}

	private int estimateHorizontalFieldRows(Class<? extends ExcelBaseModel> modelClass, ExcelWrite excelWrite) {
		ExcelCacheModel cacheModel = ExcelMetadataRegistry.getExcelCacheMapValue(modelClass);
		int fieldRows = Math.max(1, cacheModel.getFieldModelList().size());
		int rowSpan = 1 + Math.max(excelWrite.mergeTitleRowNum(), excelWrite.mergeContentRowNum());
		return Math.max(1, fieldRows * rowSpan);
	}

	private ExcelSuffixEnum resolveSuffix(String name) {
		if (name != null && name.toLowerCase().endsWith(ExcelSuffixEnum.CSV.getSuffix())) {
			return ExcelSuffixEnum.CSV;
		}
		return name != null && name.toLowerCase().endsWith(ExcelSuffixEnum.XLS.getSuffix())
				? ExcelSuffixEnum.XLS
				: ExcelSuffixEnum.XLSX;
	}
}
