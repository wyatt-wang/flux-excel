package com.github.excel.read;

import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.engine.ExcelRuntimeOptions;
import com.github.excel.helper.ExcelValidationHelper;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.read.pipeline.ExcelReadContext;
import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.read.format.ExcelReaderFormatManager;
import com.github.excel.read.handler.row.ExcelReaderRowParser;
import com.github.excel.read.format.ExcelReaderDataFormat;
import com.github.excel.util.IOUtils;
import com.github.excel.util.StringUtil;
import com.github.excel.param.ExcelReaderStreamParam;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

public class ExcelReadKernel<T extends ExcelBaseModel> {

	private final ExcelReaderFormatManager readerFormatManager;
	private final ExcelReaderTemplateValidator templateValidator;
	private final ExcelReaderPictureExecutor pictureExecutor;
	private final ExcelReaderRowParser<T> rowParser;
	private final ExcelReaderDataFormat csvDataFormat;

	public ExcelReadKernel() {
		this(ExcelRuntimeOptions.defaults());
	}

	public ExcelReadKernel(ExcelRuntimeOptions runtimeOptions) {
		this.readerFormatManager = runtimeOptions.createReaderFormatManager();
		this.templateValidator = runtimeOptions.createTemplateValidator();
		this.pictureExecutor = runtimeOptions.createPictureExecutor();
		this.rowParser = runtimeOptions.createRowParser();
		this.csvDataFormat = runtimeOptions.createCsvDataFormat();
	}

	public void createWorkbook(ExcelReadContext<T> context) {
		Workbook workbook = WorkbookHelper.createReadWorkBook(context.getReaderContext().getReaderParam());
		context.setWorkbook(workbook);
		context.setFormulaEvaluator(workbook.getCreationHelper().createFormulaEvaluator());
	}

	public void prepareRuntime(ExcelReadContext<T> context) {
		context.getReaderContext().getParserContext().setTemplateValidator(templateValidator);
		context.getReaderContext().getParserContext().setReaderFormatManager(readerFormatManager);
		context.getReaderContext().getParserContext().setPictureExecutor(pictureExecutor);
		readerFormatManager.setFormulaEvaluator(context.getFormulaEvaluator());
		Map<Integer, List<ExcelReaderModelContext<T>>> sheetModelMap = context.getReaderContext()
				.getModelMap()
				.values()
				.stream()
				.collect(Collectors.groupingBy(model -> model.getParam().getSheetIndex()));
		context.setSheetModelMap(sheetModelMap);
	}

	public void validateTemplate(ExcelReadContext<T> context) {
		String template = context.getReaderContext().getReaderParam().getTemplate();
		if (StringUtil.notEmpty(template)) {
			templateValidator.validateTemplate(template, context.getWorkbook());
		}
	}

	public void parseSheets(ExcelReadContext<T> context) {
		for (Map.Entry<Integer, List<ExcelReaderModelContext<T>>> readModelEntry : context.getSheetModelMap().entrySet()) {
			Sheet sheet = context.getWorkbook().getSheetAt(readModelEntry.getKey());
			if (Objects.isNull(sheet)) {
				continue;
			}
			Map<String, List<ExcelReaderPictureModel>> sheetPictureMap = null;
			if (context.getReaderContext().getReaderParam().getReadPicture()) {
				sheetPictureMap = pictureExecutor.getSheetPictureMap(sheet);
			}
			Map<String, Object> mergedCellValueMap = buildMergedCellValueMap(sheet, context.getFormulaEvaluator());
			fillMergedCells(sheet, mergedCellValueMap);
			parseRows(context, readModelEntry.getValue(), sheet, sheetPictureMap, mergedCellValueMap);
		}
	}

	public void applyCustomReader(ExcelReadContext<T> context) {
		if (Objects.nonNull(context.getReaderContext().getExcelCustomReader())) {
			context.getReaderContext().getExcelCustomReader().read(context.getWorkbook());
		}
	}

	public void validateModels(ExcelReadContext<T> context) {
		List<ExcelReaderModelContext<T>> fillModels = context.getReaderContext()
				.getModelMap()
				.values()
				.stream()
				.filter(model -> Objects.nonNull(model.getModel()))
				.collect(Collectors.toList());
		ExcelValidationHelper.checkTitle(fillModels);
		fillModels.forEach(readModel -> {
			ExcelValidationHelper.resetBeanColAddress(readModel.getModel(), readModel.getListTitleConfig());
			ExcelValidationHelper.validationBean(readModel, readModel.getModel());
			if (Objects.nonNull(readModel.getParam().getRowHandler())) {
				readModel.getParam().getRowHandler().handler(readModel.getModel());
			}
		});
	}

	public void flushBatch(ExcelReadContext<T> context) {
		context.getReaderContext().getModelMap().values()
				.stream()
				.filter(readModel -> readModel.getParam() instanceof ExcelReaderListParam)
				.forEach(readModel -> {
					ExcelReaderListParam<T> listParam = (ExcelReaderListParam<T>) readModel.getParam();
					if (listParam.getBatchProcess() != null
							&& readModel.getModelList() != null
							&& !readModel.getModelList().isEmpty()) {
						listParam.getBatchProcess().process(readModel.getModelList());
						readModel.getModelList().clear();
					}
				});
	}

	public void clearCoordinates(ExcelReadContext<T> context) {
		context.getReaderContext().getModelMap().values().forEach(readModel -> {
			if (readModel.getModel() != null) {
				readModel.getModel().setModelColAddress(null);
			}
			if (CollectionUtils.isNotEmpty(readModel.getModelList())) {
				readModel.getModelList().forEach(model -> model.setModelColAddress(null));
			}
		});
	}

	public void cleanup(ExcelReadContext<T> context) {
		IOUtils.closeQuietlyMulti(context.getWorkbook());
		if (context.getReaderContext().getReaderParam() instanceof ExcelReaderStreamParam
				&& context.getReaderContext().getReaderParam().getCloseInputStream()) {
			ExcelReaderStreamParam streamParam = (ExcelReaderStreamParam) context.getReaderContext().getReaderParam();
			IOUtils.closeQuietlyMulti(streamParam.getStream());
		}
	}

	private void parseRows(ExcelReadContext<T> context, List<ExcelReaderModelContext<T>> readModelList,
						   Sheet sheet, Map<String, List<ExcelReaderPictureModel>> sheetPictureMap,
						   Map<String, Object> mergedCellValueMap) {
		for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (Objects.isNull(row) || rowBlank(row)) {
				continue;
			}
			context.getReaderContext().getParserContext().setSheetPictureMap(sheetPictureMap);
			context.getReaderContext().getParserContext().setMergedCellValueMap(mergedCellValueMap);
			context.getReaderContext().getParserContext().setReadModelList(readModelList);
			context.getReaderContext().getParserContext().setRow(row);
			rowParser.rowParser(context.getReaderContext());
		}
	}

	private boolean rowBlank(Row row) {
		for (Cell cell : row) {
			if (cell.getCellType() != CellType.BLANK) {
				return false;
			}
		}
		return true;
	}

	private Map<String, Object> buildMergedCellValueMap(Sheet sheet, FormulaEvaluator formulaEvaluator) {
		Map<String, Object> valueMap = new java.util.HashMap<>();
		DataFormatter formatter = new DataFormatter();
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			org.apache.poi.ss.util.CellRangeAddress range = sheet.getMergedRegion(i);
			Row firstRow = sheet.getRow(range.getFirstRow());
			Cell firstCell = firstRow == null ? null : firstRow.getCell(range.getFirstColumn());
			Object value = formatter.formatCellValue(firstCell, formulaEvaluator);
			for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
				for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
					valueMap.put(rowIndex + ":" + colIndex, value);
				}
			}
		}
		return valueMap;
	}

	private void fillMergedCells(Sheet sheet, Map<String, Object> mergedCellValueMap) {
		if (mergedCellValueMap.isEmpty()) {
			return;
		}
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			org.apache.poi.ss.util.CellRangeAddress range = sheet.getMergedRegion(i);
			Object value = mergedCellValueMap.get(range.getFirstRow() + ":" + range.getFirstColumn());
			if (StringUtil.isEmpty(value)) {
				continue;
			}
			for (int rowIndex = range.getFirstRow(); rowIndex <= range.getLastRow(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					row = sheet.createRow(rowIndex);
				}
				for (int colIndex = range.getFirstColumn(); colIndex <= range.getLastColumn(); colIndex++) {
					if (rowIndex == range.getFirstRow() && colIndex == range.getFirstColumn()) {
						continue;
					}
					Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if (cell.getCellType() == CellType.BLANK) {
						cell.setCellValue(String.valueOf(value));
					}
				}
			}
		}
	}
}
