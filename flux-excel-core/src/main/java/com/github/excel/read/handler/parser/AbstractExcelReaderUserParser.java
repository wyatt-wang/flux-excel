package com.github.excel.read.handler.parser;

import com.github.excel.context.ExcelReaderContext;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.helper.ExcelValidationHelper;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.github.excel.read.executor.impl.ExcelReaderPictureStanderExecutor;
import com.github.excel.read.executor.impl.ExcelReaderTemplateStanderValidator;
import com.github.excel.read.handler.row.ExcelReaderRowParser;
import com.github.excel.read.executor.ExcelReaderTemplateValidator;
import com.github.excel.read.format.ExcelReaderFormatManager;
import com.github.excel.read.handler.row.ExcelReaderRowParserImpl;
import com.github.excel.util.IOUtils;
import com.github.excel.util.StringUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 读取抽象类
 */
@Slf4j
public abstract class AbstractExcelReaderUserParser<T extends ExcelBaseModel> implements ExcelReaderUserParser {
	/**
	 * 格式化缓存
	 */
	protected ExcelReaderFormatManager readerFormatManager = new ExcelReaderFormatManager();
	/**
	 * 模版校验
	 */
	protected ExcelReaderTemplateValidator templateValidator = new ExcelReaderTemplateStanderValidator();
	/**
	 * 图片获取
	 */
	protected ExcelReaderPictureExecutor pictureExecutor = new ExcelReaderPictureStanderExecutor();
	/**
	 * 行解析
	 */
	protected static final ExcelReaderRowParser EXCEL_READER_ROW_PARSER;
	/**
	 * 上下文参数
	 */
	private ExcelReaderContext<T> readerContext ;

	static {
		ServiceLoader<ExcelReaderRowParser> serviceLoader = ServiceLoader.load(ExcelReaderRowParser.class);
		List<ExcelReaderRowParser> rowParsers = new ArrayList<>();
		for (ExcelReaderRowParser rowParser : serviceLoader) {
			rowParsers.add(rowParser);
		}
		EXCEL_READER_ROW_PARSER = rowParsers.stream().filter(e -> !(e instanceof ExcelReaderRowParserImpl)).findFirst().orElse(new ExcelReaderRowParserImpl());
	}
	public AbstractExcelReaderUserParser(ExcelReaderContext<T> readerContext) {
		readerContext.getParserContext().setTemplateValidator(templateValidator);
		readerContext.getParserContext().setReaderFormatManager(readerFormatManager);
		readerContext.getParserContext().setPictureExecutor(pictureExecutor);
		this.readerContext = readerContext;
	}

	public void parseExcel(){
		// todo resetBeanColAddress 调用一下
		// 对 sheetName / sheetIndex 分组，避免指定 sheetName 时和默认 sheetIndex=0 串组
		Map<String, List<ExcelReaderModelContext<T>>> sheetModelMap = readerContext.getModelMap().values().stream()
				.collect(Collectors.groupingBy(this::resolveSheetGroupKey));
		Workbook workbook = WorkbookHelper.createReadWorkBook(readerContext.getReaderParam());
		try {
			FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
			readerFormatManager.setFormulaEvaluator(formulaEvaluator);
			// 1. 校验模版
			if(StringUtil.notEmpty(readerContext.getReaderParam().getTemplate())) {
				templateValidator.validateTemplate(readerContext.getReaderParam().getTemplate(), workbook);
			}
			// 2、根据sheetIndex进行循环，填充里面的对象
			for (Map.Entry<String, List<ExcelReaderModelContext<T>>> readModelEntry : sheetModelMap.entrySet()) {
				List<ExcelReaderModelContext<T>> readModelList = readModelEntry.getValue();
				Sheet sheet = resolveSheet(workbook, readModelList);
				if (Objects.isNull(sheet)) {
					continue;
				}
				Map<String, List<ExcelReaderPictureModel>> sheetPictureMap = null;
				// 3、读取Picture
				if (readerContext.getReaderParam().getReadPicture()) {
					sheetPictureMap = pictureExecutor.getSheetPictureMap(sheet);
				}
				// 4.循环需要读取Excel Sheet 下面的所有行
				int startRow = resolveStartRow(sheet, readModelList);
				int endRow = resolveEndRow(sheet, readModelList);
				int readRows = 0;
				Integer maxRows = resolveMaxRows(readModelList);
				for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
					if (Objects.nonNull(maxRows) && readRows >= maxRows) {
						break;
					}
					Row row = sheet.getRow(rowIndex);
					if (Objects.isNull(row)) {
						continue;
					}
					// 判断是否是空白单元格行
					boolean rowBlank = true;
					for (Cell cell:row){
						if(cell.getCellType() != CellType.BLANK){
							rowBlank = false;
							break;
						}
					}
					if(!rowBlank || !Boolean.TRUE.equals(readerContext.getReaderParam().getIgnoreEmptyRow())) {
						readerContext.getParserContext().setSheetPictureMap(sheetPictureMap);
						readerContext.getParserContext().setReadModelList(readModelList);
						readerContext.getParserContext().setRow(row);
						EXCEL_READER_ROW_PARSER.rowParser(readerContext);
						readRows++;
					}
				}
			}
			// 5. 执行回调
			if(Objects.nonNull(readerContext.getExcelCustomReader())) {
				readerContext.getExcelCustomReader().read(workbook);
			}
		} finally {
			IOUtils.closeQuietlyMulti(workbook);
			if (readerContext.getReaderParam() instanceof ExcelReaderStreamParam) {
				ExcelReaderStreamParam streamParam = (ExcelReaderStreamParam) readerContext.getReaderParam();
				if (readerContext.getReaderParam().getCloseInputStream()) {
					IOUtils.closeQuietlyMulti(streamParam.getStream());
				}
			}
		}
		// 2. 执行校验Bean
		List<ExcelReaderModelContext<T>> fillModes = readerContext.getModelMap().values().stream().filter(e -> Objects.nonNull(e.getModel())).collect(Collectors.toList());
		ExcelValidationHelper.checkTitle(fillModes);
		fillModes.forEach(readeModel ->{
			// 2.1 设置list title坐标
			ExcelValidationHelper.resetBeanColAddress(readeModel.getModel(), readeModel.getListTitleConfig());
			// 2.2 进行校验
			ExcelValidationHelper.validationBean(readeModel, readeModel.getModel());
			if (Objects.nonNull(readeModel.getParam().getRowHandler())) {
				// 2.3 行处理器调用
				readeModel.getParam().getRowHandler().handler(readeModel.getModel());
			}
		});
		// 3. list 回调
		readerContext.getModelMap().values().stream().filter(e -> e.getParam() instanceof ExcelReaderListParam).forEach(readModelDto -> {
			ExcelReaderListParam<T> listParam = (ExcelReaderListParam<T>)readModelDto.getParam();
			// batch processing
			if (listParam.getBatchProcess() != null &&
					readModelDto.getModelList() != null && !readModelDto.getModelList().isEmpty()) {
				// 批处理，处理最后遗留数据
				listParam.getBatchProcess().process(readModelDto.getModelList());
				readModelDto.getModelList().clear();
			}
		});
		// 4. 清空坐标
		readerContext.getModelMap().values().forEach(e->{
			if(e.getModel()!=null){
				e.getModel().setModelColAddress(null);
			}
			if(CollectionUtils.isNotEmpty(e.getModelList())){
				e.getModelList().forEach(ele -> ele.setModelColAddress(null));
			}
		});
	}

	private String resolveSheetGroupKey(ExcelReaderModelContext<T> modelContext) {
		String sheetName = modelContext.getParam().getSheetName();
		if (StringUtil.notEmpty(sheetName)) {
			return "name:" + sheetName;
		}
		return "index:" + modelContext.getParam().getSheetIndex();
	}

	private Sheet resolveSheet(Workbook workbook, List<ExcelReaderModelContext<T>> readModelList) {
		Optional<String> sheetName = readModelList.stream()
				.map(model -> model.getParam().getSheetName())
				.filter(StringUtil::notEmpty)
				.findFirst();
		if (sheetName.isPresent()) {
			return workbook.getSheet(sheetName.get());
		}
		Integer sheetIndex = readModelList.get(0).getParam().getSheetIndex();
		return workbook.getSheetAt(sheetIndex);
	}

	private int resolveStartRow(Sheet sheet, List<ExcelReaderModelContext<T>> readModelList) {
		return readModelList.stream()
				.map(model -> model.getParam().getDataStartRow())
				.filter(Objects::nonNull)
				.min(Integer::compareTo)
				.orElseGet(() -> readModelList.stream()
						.map(model -> model.getParam().getHeadRowNumber())
						.filter(Objects::nonNull)
						.max(Integer::compareTo)
						.orElse(sheet.getFirstRowNum()));
	}

	private int resolveEndRow(Sheet sheet, List<ExcelReaderModelContext<T>> readModelList) {
		return readModelList.stream()
				.map(model -> model.getParam().getDataEndRow())
				.filter(Objects::nonNull)
				.max(Integer::compareTo)
				.orElse(sheet.getLastRowNum());
	}

	private Integer resolveMaxRows(List<ExcelReaderModelContext<T>> readModelList) {
		return readModelList.stream()
				.map(model -> model.getParam().getMaxRows())
				.filter(Objects::nonNull)
				.max(Integer::compareTo)
				.orElse(null);
	}

}
