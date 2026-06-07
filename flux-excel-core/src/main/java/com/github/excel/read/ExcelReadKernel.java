package com.github.excel.read;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.helper.ExcelValidationHelper;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.read.pipeline.ExcelReadContext;
import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.read.executor.impl.ExcelReaderPictureStanderExecutor;
import com.github.excel.read.executor.impl.ExcelReaderTemplateStanderValidator;
import com.github.excel.read.format.ExcelReaderFormatManager;
import com.github.excel.read.handler.row.ExcelReaderRowParser;
import com.github.excel.read.handler.row.ExcelReaderRowParserImpl;
import com.github.excel.read.format.ExcelReaderDataFormat;
import com.github.excel.read.format.ExcelDefaultReaderDataFormat;
import com.github.excel.util.IOUtils;
import com.github.excel.util.StringUtil;
import com.github.excel.param.ExcelReaderStreamParam;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.ServiceLoader;
import java.util.stream.Collectors;

public class ExcelReadKernel<T extends ExcelBaseModel> {

	private final ExcelReaderFormatManager readerFormatManager = new ExcelReaderFormatManager();
	private final ExcelReaderTemplateValidator templateValidator = new ExcelReaderTemplateStanderValidator();
	private final ExcelReaderPictureExecutor pictureExecutor = new ExcelReaderPictureStanderExecutor();
	private final ExcelReaderRowParser<T> rowParser = loadRowParser();
	private final ExcelReaderDataFormat csvDataFormat = new ExcelDefaultReaderDataFormat();

	public ExcelReadKernel() {
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
			parseRows(context, readModelEntry.getValue(), sheet, sheetPictureMap);
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

	@SuppressWarnings("unchecked")
	private ExcelReaderRowParser<T> loadRowParser() {
		ServiceLoader<ExcelReaderRowParser> serviceLoader = ServiceLoader.load(ExcelReaderRowParser.class);
		List<ExcelReaderRowParser> rowParsers = new ArrayList<>();
		for (ExcelReaderRowParser rowParser : serviceLoader) {
			rowParsers.add(rowParser);
		}
		return (ExcelReaderRowParser<T>) rowParsers.stream()
				.filter(rowParser -> !(rowParser instanceof ExcelReaderRowParserImpl))
				.findFirst()
				.orElse(new ExcelReaderRowParserImpl<>());
	}

	private void parseRows(ExcelReadContext<T> context, List<ExcelReaderModelContext<T>> readModelList,
						   Sheet sheet, Map<String, List<ExcelReaderPictureModel>> sheetPictureMap) {
		for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (Objects.isNull(row) || rowBlank(row)) {
				continue;
			}
			context.getReaderContext().getParserContext().setSheetPictureMap(sheetPictureMap);
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
}
