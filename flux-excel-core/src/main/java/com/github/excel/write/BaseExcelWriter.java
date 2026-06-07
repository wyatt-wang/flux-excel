package com.github.excel.write;

import com.github.excel.annotation.ExcelWrite;
import com.github.excel.annotation.ExcelWriteProperty;
import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.constant.ExcelErrorMsgConstant;
import com.github.excel.enums.*;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.*;
import com.github.excel.param.*;
import com.github.excel.util.ReflectCacheUtil;
import com.github.excel.util.StringUtil;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.write.style.ExcelBasicStyle;
import com.google.common.base.Strings;
import com.google.common.base.Throwables;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 通用服务
 */
@Slf4j
public class BaseExcelWriter {

	/**
	 * 生成Excel数据，list模型
	 */
	protected List<ExcelWriterModel> exportModelList = new ArrayList<>();
	/**
	 * 生成Excel数据，单个bean模型
	 */
	protected List<ExcelWriterModel> exportBeanList = new ArrayList<>();
	/**
	 * 自定义填充单元格List
	 */
	protected List<ExcelWriterCellParam> customColumnModelList = new ArrayList<>();
	/**
	 * 自定义Merge单元格List
	 */
	protected List<ExcelWriterMergeParam> mergeCustomColumnModelList = new ArrayList<>();
	/**
	 * 自定义样式
	 */
	protected List<Class<? extends AbstractExcelStyle>> styleList = new ArrayList<>(ExcelConstant.TOW_INT);
	/**
	 * 样式缓存
	 */
	protected ThreadLocal<Map<String, CellStyle>> styleLocal = ThreadLocal.withInitial(() -> Maps.newHashMap());

	/**
	 * 字体缓存
	 */
	protected ThreadLocal<Map<String, Font>> fontLocal = ThreadLocal.withInitial(() -> Maps.newHashMap());
	/**
	 * 颜色缓存
	 */
	protected ThreadLocal<Map<String, Color>> colorLocal = ThreadLocal.withInitial(() -> Maps.newHashMap());
	/**
	 * 字段格式化器缓存
	 */
	protected Map<Class<? extends ExcelWriterDataFormat>, ExcelWriterDataFormat> dataFormatMap = new HashMap<>();
	/**
	 * Excel 名称
	 */
	protected String template;
	/**
	 * 自定义导出器
	 */
	protected ExcelCustomWriter customWrite;
	/**
	 * 默认格式化器
	 */
	protected ExcelWriterDataFormat dataFormat = new ExcelDefaultWriterDataFormat();
	/**
	 * class 信息缓存
	 */
	protected ReflectCacheUtil reflectCacheUtil = new ReflectCacheUtil();
	/**
	 * 是否启用流模式
	 */
	protected boolean streaming = false;
	/**
	 * 排除字段map
	 */
	protected Map<Class<? extends ExcelBaseModel>, Map<String, String>> excludeFieldMap = new HashMap<>();
	/**
	 * 校验或批注map
	 */
	protected Map<Class<? extends ExcelBaseModel>, Map<String, ExcelWriterCommentParam>> commentMap = new HashMap<>();
	/**
	 * 字段条件样式map
	 */
	protected Map<Class<? extends ExcelBaseModel>, Map<String, List<ExcelWriterConditionalStyleParam>>> conditionalStyleMap = new HashMap<>();
	/**
	 * 自定义范围条件样式
	 */
	protected List<ExcelWriterConditionalStyleParam> conditionalStyleList = new ArrayList<>();
	protected Map<String, String> runtimeHeaderAliases = new LinkedHashMap<>();
	protected Set<String> runtimeIncludeFields = new LinkedHashSet<>();
	protected boolean runtimeOnlyAlias = false;
	protected Map<Class<?>, java.util.function.Function<Object, Object>> runtimeTypeConverters = new LinkedHashMap<>();
	protected Map<String, java.util.function.Function<Object, Object>> runtimeFieldConverters = new LinkedHashMap<>();
	/**
	 * 设置选中sheet
	 */
	protected String selectSheet = null;
	/**
	 * 设置选中sheet
	 */
	protected Class<? extends ExcelBaseModel> listCla = null;

	/**
	 * 填充没有数据提示
	 */
	protected boolean noneDataTips = true;

	/**
	 * 自增map
	 */
	protected Map<Class<? extends ExcelBaseModel>, AtomicLong> incrementSeqMap = Maps.newConcurrentMap();

	/**
	 * sheet 最大行数
	 */
	protected int sheetRowMaxCount ;




	/**
	 * 导出到文件
	 *
	 * @param outputStream 导出流
	 */
	protected void writeToNewFile(OutputStream outputStream, String excelName) {
		SXSSFWorkbook sxssfWorkbook = null;
		ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookThreadLocal;
			if (excelName.endsWith(ExcelConstant.XLSX_STR)) {
				if (streaming) {
					workbookThreadLocal = WorkbookCachePool.addBasicStyle(WorkbookHelper.createStreamingXlsxWorkBook(
							ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE,
							false,
							false
					));
				} else {
					workbookThreadLocal = WorkbookCachePool.addBasicStyle(new org.apache.poi.xssf.usermodel.XSSFWorkbook());
				}
		} else {
			workbookThreadLocal = WorkbookCachePool.addBasicStyle(new HSSFWorkbook());
		}

		if (Objects.isNull(workbookThreadLocal)) {
			throw new ExcelWriterException("Failed to fetch workbook from cache");
		}
		WorkbookCachePool.WorkbookCacheModel cacheModel = workbookThreadLocal.get();
		styleLocal.set(cacheModel.getStyleMap());
		fontLocal.set(cacheModel.getFontMap());
		colorLocal.set(cacheModel.getColorMap());

		try (Workbook workbook = cacheModel.getWorkbook()) {

			if (workbook instanceof SXSSFWorkbook) {
				sxssfWorkbook = (SXSSFWorkbook) workbook;
			}
			// 初始化样式
			initStyle(workbook);
			CreationHelper createHelper = workbook.getCreationHelper();

			fillCustomColumn(workbook, createHelper);
			fillMergeCustomColumn(workbook, createHelper);
			for (ExcelWriterModel exportModel : exportBeanList) {
				Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
				fillBean(workbook, createHelper, exportModel, sheet);
			}
			for (ExcelWriterModel exportModel : exportModelList) {
				Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, exportModel.getSheetName());
				fillBeanList(workbook, createHelper, exportModel, sheet);
			}

			if (Objects.nonNull(customWrite)) {
				customWrite.execute(workbook);
			}
			addNoResultData(workbook, createHelper);
			selectSheet(workbook);
			workbook.write(outputStream);
		} catch (IllegalAccessException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		} catch (IOException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		} catch (InvocationTargetException e) {
			log.error("Export excel failed, cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		} finally {
			if (Objects.nonNull(sxssfWorkbook)) {
				sxssfWorkbook.dispose();
			}
			styleLocal.remove();
			fontLocal.remove();
		}
	}

	/**
	 * 设置sheet为选中状态
	 *
	 * @param workbook
	 */
	protected void selectSheet(Workbook workbook) {
		if (StringUtil.notEmpty(selectSheet)) {
			Sheet sheet = workbook.getSheet(selectSheet);
			if (Objects.nonNull(sheet)) {
				sheet.setSelected(true);
			}
		}
	}




	/**
	 * 初始化样式
	 *
	 * @param workbook workBook
	 */
	protected void initStyle(Workbook workbook) {
		try {
			for (Class<? extends AbstractExcelStyle> styleClass : styleList) {
				initStyle(workbook, styleClass);
			}
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException("Init style error");
		}
	}

	/**
	 * 初始化样式
	 *
	 * @param workbook   workBook
	 * @param styleClass styleClass
	 */
	protected void initStyle(Workbook workbook, Class<? extends AbstractExcelStyle> styleClass) {
		try {

			Constructor<? extends AbstractExcelStyle> constructor = styleClass.getConstructor(styleClass.getConstructors()[ExcelConstant.ZERO_SHORT].getParameterTypes());
			AbstractExcelStyle excelStyle = constructor.newInstance(workbook, styleLocal.get(), fontLocal.get(), colorLocal.get());
			excelStyle.addNewFont();
			excelStyle.addNewStyle();
			excelStyle.addNewColor();
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException("Init style error");
		}
	}

	/**
	 * 获取格式化器
	 *
	 * @param formatClass 格式化器class
	 */
	protected ExcelWriterDataFormat getFormatter(Class<? extends ExcelWriterDataFormat> formatClass) {
		if (formatClass == ExcelDefaultWriterDataFormat.class) {
			return null;
		}
		ExcelWriterDataFormat format = dataFormatMap.get(formatClass);
		if (null == format) {
			try {
				format = formatClass.newInstance();
			} catch (Exception e) {
				log.error("Error by create formatter , cause:{}", Throwables.getStackTraceAsString(e));
			}
			dataFormatMap.put(formatClass, format);
		}
		return format;
	}

	/**
	 * 填充自定义列
	 *
	 * @param workbook
	 * @param createHelper
	 */
	protected void fillCustomColumn(Workbook workbook, CreationHelper createHelper) {
		for (ExcelWriterCellParam columnModel : customColumnModelList) {
			fillColumn(workbook, createHelper, columnModel, columnModel.getRowIndex(), columnModel.getColIndex());
		}
	}

	/**
	 * 填充column
	 *
	 * @param workbook
	 * @param createHelper
	 * @param columnModel
	 * @param rowIndex
	 * @param colIndex
	 */
	protected void fillColumn(Workbook workbook, CreationHelper createHelper, ExcelWriterCellParam columnModel, int rowIndex, int colIndex) {
		Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, columnModel.getSheetName());
		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, columnModel.getColWidth());
		ExcelHelper.setRowHeight(row, columnModel.getRowHeight());
		ExcelWriterDataFormat formatter = getFormatter(columnModel.getDataFormat());
		Object value = formatValue(columnModel.getValue(), columnModel.getFormatPattern(), formatter);

		if (columnModel.getFillType() == ExcelWriterColumnFillTypeEnum.APPEND) {
			String cellValue = cell.getStringCellValue();
			if (StringUtil.notEmpty(cellValue)) {
				value = cellValue + value;
			}
		}
		setCellValueAndStyle(cell, value, columnModel.getStyleName(), null, createHelper);
	}

	/**
	 * 填充合并列
	 *
	 * @param workbook
	 * @param createHelper
	 */
	protected void fillMergeCustomColumn(Workbook workbook, CreationHelper createHelper) {
		for (ExcelWriterMergeParam columnModel : mergeCustomColumnModelList) {
			Sheet sheet = ExcelHelper.getSheetOrCreate(workbook, columnModel.getSheetName());
			sheet.addMergedRegion(new CellRangeAddress(columnModel.getRowIndex(), columnModel.getEndRowIndex(), columnModel.getColIndex(), columnModel.getEndColIndex()));

			fillColumn(workbook, createHelper, columnModel, columnModel.getRowIndex(), columnModel.getColIndex());
		}
	}

	/**
	 * 填充单个bean
	 *
	 * @param workbook     workbook
	 * @param createHelper 创建对象的
	 * @param exportModel  数据组装实体
	 * @param sheet        sheet
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void fillBean(Workbook workbook, CreationHelper createHelper, ExcelWriterModel exportModel, Sheet sheet) throws IllegalAccessException, InvocationTargetException {
		ExcelCacheModel cacheModel = exportModel.getCacheModel();
		ExcelWrite excelWrite = cacheModel.getExcelWrite();
		int colIndex = excelWrite.colIndex(), rowIndex = excelWrite.rowIndex(), initColIndex = excelWrite.colIndex();
		if (null != exportModel.getRowIndex()) {
			rowIndex = exportModel.getRowIndex();
		}

		if (null != exportModel.getColIndex()) {
			colIndex = exportModel.getColIndex();
			initColIndex = exportModel.getColIndex();
		}
		// todo
		if (null != exportModel.getRowIndex()) {
			rowIndex = exportModel.getRowIndex();
		}

		doFillBean(workbook, createHelper, exportModel, sheet, cacheModel, excelWrite, colIndex, rowIndex, initColIndex);
	}

	/**
	 * 执行填充单个对象
	 *
	 * @param workbook
	 * @param createHelper
	 * @param exportModel
	 * @param sheet
	 * @param cacheModel
	 * @param excelWrite
	 * @param colIndex
	 * @param rowIndex
	 * @param initColIndex
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void doFillBean(Workbook workbook, CreationHelper createHelper, ExcelWriterModel exportModel, Sheet sheet, ExcelCacheModel cacheModel, ExcelWrite excelWrite, int colIndex, int rowIndex, final int initColIndex) throws IllegalAccessException, InvocationTargetException {
		boolean hasCycle = false;
		ExcelWriterFillStyleEnum lasFileStyle = null;
		List<Integer> rowIndexList = Lists.newArrayList();
		Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
		for (ExcelCacheFieldModel cacheFieldModel : cacheModel.getFieldModelList()) {
			boolean isMap = cacheFieldModel.isMap();
			ExcelWriteProperty exportCell = cacheFieldModel.getExportCell();
			if (Objects.isNull(exportCell)) {
				continue;
			}
			if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(cacheFieldModel.getFieldName()))) {
				continue;
			}
			if (skipRuntimeField(cacheFieldModel)) {
				continue;
			}
			ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());

			Object value = cacheFieldModel.getGetMethod().invoke(exportModel.getDataModel());
			value = convertRuntimeValue(cacheFieldModel, value);
			String linkName = resolveLinkName(cacheFieldModel, exportModel.getDataModel());
			value = formatValue(value, exportCell.formatPattern(), formatter);

			// 填充指定的列
			if (exportCell.rowIndex() != ExcelConstant.MINUS_ONE_SHORT && exportCell.colIndex() != ExcelConstant.MINUS_ONE_SHORT) {
				Row row = ExcelHelper.getRowOrCreate(sheet, exportCell.rowIndex());
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				fillContent(workbook, sheet, exportCell.colIndex(), rowIndex, exportCell, value, row, createHelper, excelWrite, isMap, cacheFieldModel, linkName);
			} else {
				// 自动填充
				if (hasCycle) {
					if (ExcelWriterFillStyleEnum.HORIZONTAL == exportCell.fillStyle()) {
						colIndex++;
						if (exportCell.verticalNewLine() && ExcelWriterFillStyleEnum.VERTICAL == lasFileStyle) {
							rowIndex++;
							colIndex = initColIndex;
						}
					} else if (ExcelWriterFillStyleEnum.VERTICAL == exportCell.fillStyle()) {
						if (rowIndexList.size() > ExcelConstant.ZERO_SHORT) {
							rowIndex = rowIndexList.stream().max(Comparator.comparing(Integer::intValue)).get();
							rowIndexList.clear();
						}
						if(excelWrite.fillContent()) {
							rowIndex++;
						}
						// 判断是否填充标题
						if(!excelWrite.fillContent() && exportCell.colIndex() != ExcelConstant.MINUS_ONE_SHORT){
							colIndex = exportCell.colIndex();
							rowIndex++;
						}else {
							colIndex = initColIndex;
						}
					}
				}
				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				hasCycle = true;
				lasFileStyle = exportCell.fillStyle();

				ColPointDto pointDto = fillContent(workbook, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelWrite, isMap, cacheFieldModel, linkName);
				colIndex = pointDto.colIndex;
				if(!excelWrite.fillContent() && exportCell.colIndex() != ExcelConstant.MINUS_ONE_SHORT){
					colIndex = exportCell.colIndex();
				}
				if (ExcelWriterFillStyleEnum.HORIZONTAL == exportCell.fillStyle()) {
					rowIndexList.add(pointDto.rowIndex);
				} else if (ExcelWriterFillStyleEnum.VERTICAL == exportCell.fillStyle()) {
					rowIndex = pointDto.rowIndex;
				}
			}
		}
	}

	/**
	 * 填充列表
	 *
	 * @param workbook     workbook
	 * @param createHelper 创建对象的
	 * @param exportModel  数据组装实体
	 * @param sheet        sheet
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 */
	protected void fillBeanList(Workbook workbook, CreationHelper createHelper, ExcelWriterModel exportModel, Sheet sheet) {
		ExcelCacheModel cacheModel = exportModel.getCacheModel();
		int rowIndex = cacheModel.getExcelWrite().rowIndex(), colIndex = cacheModel.getExcelWrite().colIndex(), initColIndex = cacheModel.getExcelWrite().colIndex(), initRowIndex = cacheModel.getExcelWrite().rowIndex();
		ExcelWriterFillStyleEnum fillStyleEnum = exportModel.getCacheModel().getExcelWrite().fillStyle();
		if (null != exportModel.getRowIndex()) {
			rowIndex = exportModel.getRowIndex();
			initRowIndex = exportModel.getRowIndex();
		}

		if (null != exportModel.getColIndex()) {
			colIndex = exportModel.getColIndex();
			initColIndex = exportModel.getColIndex();
		}
		// 移动行
		if (cacheModel.getExcelWrite().fillType() == ExcelWriterListFillTypeEnum.SHIFT) {
			ExcelHelper.shiftRows(sheet, rowIndex, cacheModel.getExcelWrite().fillTitle() ? exportModel.getDataModelList().size() + ExcelConstant.ONE_INT : exportModel.getDataModelList().size());
		}
		List<ExcelCacheFieldModel> fillFieldModelList = new ArrayList<>();
		Map<String, String> excludeMap = excludeFieldMap.get(exportModel.getExcelModelClass());
		Map<String, Integer> commentPointIndexMap = new HashMap<>();
		int i = colIndex;
		for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
			if (Objects.nonNull(excludeMap) && Objects.nonNull(excludeMap.get(fieldModel.getFieldName()))) {
				continue;
			}
			if (skipRuntimeField(fieldModel)) {
				continue;
			}
			commentPointIndexMap.put(fieldModel.getFieldName(), i);
			fillFieldModelList.add(fieldModel);
			i++;
		}
		//填充标题
		ColPointDto colPointDto = fillListTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, initColIndex, initRowIndex, exportModel, fillFieldModelList);
		rowIndex = colPointDto.getRowIndex();
		colIndex = colPointDto.getColIndex();
		if (exportModel.getDataModelList().size() > ExcelConstant.ZERO_SHORT) {
			// 填充内容
			ColPointDto contentLastPoint = fillListContent(workbook, exportModel, cacheModel, sheet, createHelper, fillFieldModelList, rowIndex, colIndex, fillStyleEnum, initColIndex, initRowIndex);
			fillListValidationOrComment(workbook, sheet, rowIndex, contentLastPoint.getRowIndex(), commentPointIndexMap, commentMap.get(exportModel.getExcelModelClass()), exportModel);
			fillListConditionalStyles(sheet, rowIndex, contentLastPoint.getRowIndex() - ExcelConstant.ONE_INT, commentPointIndexMap,
					conditionalStyleMap.get(exportModel.getExcelModelClass()), exportModel);
		}
	}

	/**
	 * 添加校验或备注
	 *
	 * @param sheet
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param commentColIndexMap
	 * @param commentModelMap
	 */
	protected void fillListValidationOrComment(Workbook workbook, Sheet sheet, int startRowIndex, int endRowIndex, Map<String, Integer> commentColIndexMap, Map<String, ExcelWriterCommentParam> commentModelMap, ExcelWriterModel exportModel) {
		if (Objects.isNull(commentModelMap)) {
			return;
		}
		if (exportModel.getCacheModel().getExcelWrite().fillStyle() != ExcelWriterFillStyleEnum.VERTICAL) {
			return;
		}
		for (Map.Entry<String, ExcelWriterCommentParam> entry : commentModelMap.entrySet()) {
			Integer colIndex = commentColIndexMap.get(entry.getKey());
			if (Objects.isNull(colIndex)) {
				continue;
			}
			ExcelWriterCommentParam excelWriterCommentParam = entry.getValue();
			if (exportModel.getCacheModel().getExcelWrite().fillTitle() && StringUtil.notEmpty(excelWriterCommentParam.getCommentText())) {
				ExcelHelper.createComment(workbook, sheet, startRowIndex - ExcelConstant.ONE_INT, colIndex, ExcelConstant.NULL_STR, excelWriterCommentParam.getCommentText(), fontLocal.get().get(excelWriterCommentParam.getCommentFontName()));
			}
		}
	}

	protected void fillListConditionalStyles(Sheet sheet, int startRowIndex, int endRowIndex,
											 Map<String, Integer> fieldColIndexMap,
											 Map<String, List<ExcelWriterConditionalStyleParam>> fieldConditionalStyleMap,
											 ExcelWriterModel exportModel) {
		if (Objects.isNull(fieldConditionalStyleMap)) {
			return;
		}
		if (exportModel.getCacheModel().getExcelWrite().fillStyle() != ExcelWriterFillStyleEnum.VERTICAL) {
			return;
		}
		if (endRowIndex < startRowIndex) {
			return;
		}
		for (Map.Entry<String, List<ExcelWriterConditionalStyleParam>> entry : fieldConditionalStyleMap.entrySet()) {
			Integer colIndex = fieldColIndexMap.get(entry.getKey());
			if (Objects.isNull(colIndex)) {
				continue;
			}
			for (ExcelWriterConditionalStyleParam conditionalStyleParam : entry.getValue()) {
				ExcelWriterConditionalStyleParam resolvedParam = copyConditionalStyleParam(conditionalStyleParam);
				resolvedParam.setSheetName(sheet.getSheetName());
				resolvedParam.setRowIndex(startRowIndex);
				resolvedParam.setEndRowIndex(endRowIndex);
				resolvedParam.setColIndex(colIndex);
				resolvedParam.setEndColIndex(colIndex);
				applyConditionalStyle(sheet, resolvedParam);
			}
		}
	}

	protected void applyConditionalStyles(Workbook workbook) {
		if (Objects.isNull(conditionalStyleList)) {
			return;
		}
		for (ExcelWriterConditionalStyleParam conditionalStyleParam : conditionalStyleList) {
			if (conditionalStyleParam == null || StringUtil.isEmpty(conditionalStyleParam.getSheetName())) {
				continue;
			}
			Sheet sheet = workbook.getSheet(conditionalStyleParam.getSheetName());
			if (Objects.isNull(sheet)) {
				continue;
			}
			applyConditionalStyle(sheet, conditionalStyleParam);
		}
	}

	protected void applyConditionalStyle(Sheet sheet, ExcelWriterConditionalStyleParam conditionalStyleParam) {
		ConditionalFormattingRule rule = createConditionalFormattingRule(sheet, conditionalStyleParam);
		if (Objects.isNull(rule)) {
			return;
		}
		applyConditionalFormattingStyle(rule, conditionalStyleParam);
		CellRangeAddress[] regions = new CellRangeAddress[] {
				new CellRangeAddress(
						conditionalStyleParam.getRowIndex(),
						resolveEndIndex(conditionalStyleParam.getEndRowIndex(), conditionalStyleParam.getRowIndex()),
						conditionalStyleParam.getColIndex(),
						resolveEndIndex(conditionalStyleParam.getEndColIndex(), conditionalStyleParam.getColIndex())
				)
		};
		sheet.getSheetConditionalFormatting().addConditionalFormatting(regions, rule);
	}

	private ConditionalFormattingRule createConditionalFormattingRule(Sheet sheet, ExcelWriterConditionalStyleParam conditionalStyleParam) {
		SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();
		if (StringUtil.notEmpty(conditionalStyleParam.getFormula())) {
			return formatting.createConditionalFormattingRule(conditionalStyleParam.getFormula());
		}
		if (Objects.isNull(conditionalStyleParam.getOperator()) || StringUtil.isEmpty(conditionalStyleParam.getStart())) {
			return null;
		}
		if (StringUtil.notEmpty(conditionalStyleParam.getEnd())) {
			return formatting.createConditionalFormattingRule(conditionalStyleParam.getOperator().getPoiCode(),
					conditionalStyleParam.getStart(), conditionalStyleParam.getEnd());
		}
		return formatting.createConditionalFormattingRule(conditionalStyleParam.getOperator().getPoiCode(),
				conditionalStyleParam.getStart());
	}

	private void applyConditionalFormattingStyle(ConditionalFormattingRule rule, ExcelWriterConditionalStyleParam conditionalStyleParam) {
		if (Objects.nonNull(conditionalStyleParam.getFillForegroundColor())) {
			PatternFormatting patternFormatting = rule.createPatternFormatting();
			patternFormatting.setFillBackgroundColor(conditionalStyleParam.getFillForegroundColor());
			patternFormatting.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
		}
		if (Objects.nonNull(conditionalStyleParam.getFontColor())
				|| Objects.nonNull(conditionalStyleParam.getBold())
				|| Objects.nonNull(conditionalStyleParam.getItalic())) {
			FontFormatting fontFormatting = rule.createFontFormatting();
			if (Objects.nonNull(conditionalStyleParam.getFontColor())) {
				fontFormatting.setFontColorIndex(conditionalStyleParam.getFontColor());
			}
			fontFormatting.setFontStyle(Boolean.TRUE.equals(conditionalStyleParam.getItalic()),
					Boolean.TRUE.equals(conditionalStyleParam.getBold()));
		}
	}

	private int resolveEndIndex(Integer endIndex, Integer startIndex) {
		return Objects.nonNull(endIndex) ? endIndex : startIndex;
	}

	private ExcelWriterConditionalStyleParam copyConditionalStyleParam(ExcelWriterConditionalStyleParam source) {
		ExcelWriterConditionalStyleParam target = new ExcelWriterConditionalStyleParam();
		target.setSheetName(source.getSheetName());
		target.setRowIndex(source.getRowIndex());
		target.setColIndex(source.getColIndex());
		target.setFillTemplate(source.getFillTemplate());
		target.setEndRowIndex(source.getEndRowIndex());
		target.setEndColIndex(source.getEndColIndex());
		target.setFormula(source.getFormula());
		target.setOperator(source.getOperator());
		target.setStart(source.getStart());
		target.setEnd(source.getEnd());
		target.setFillForegroundColor(source.getFillForegroundColor());
		target.setFontColor(source.getFontColor());
		target.setBold(source.getBold());
		target.setItalic(source.getItalic());
		return target;
	}

	protected ColPointDto fillListMapTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, ExcelCacheFieldModel fieldModel, ExcelWriterModel exportModel) {
		try {
			ExcelBaseModel excelBaseModel = exportModel.getDataModelList().get(ExcelConstant.ZERO_SHORT);
			ExcelWriteProperty exportCell = fieldModel.getExportCell();
			Object value = fieldModel.getGetMethod().invoke(excelBaseModel);
			if (value instanceof Map) {
				Iterator keyIterator = ((Map) value).keySet().iterator();
				while (keyIterator.hasNext()) {
					String title = "";
					String titleName = resolveRuntimeTitle(fieldModel);
					if (StringUtil.notEmpty(titleName)) {
						title = title + titleName;
					}
					String key = keyIterator.next().toString();
					if (Strings.isNullOrEmpty(title)) {
						title = key;
					} else {
						title = title + ExcelConstant.SEPARATOR + key;
					}

					String styleName = StringUtil.notEmpty(exportCell.titleStyleName()) ? exportCell.titleStyleName() : cacheModel.getExcelWrite().titleStyleName();
					Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
					Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
					ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
					ExcelHelper.setRowHeight(row, fieldModel.getContentRowHeight());
					setCellValueAndStyle(cell, title, styleName, null, createHelper);

					if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
						colIndex++;
					} else {
						rowIndex++;
					}
				}
			}
		} catch (Exception e) {
			log.error(Throwables.getStackTraceAsString(e));
			//invoke方法执行出错
			throw new ExcelWriterException("method.invoke.fail");
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	protected ColPointDto fillListBeanTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, ExcelCacheFieldModel fieldModel) {
		String titleName = resolveRuntimeTitle(fieldModel);

		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
		ExcelHelper.setRowHeight(row, fieldModel.getTitleRowHeight());
		setCellValueAndStyle(cell, titleName, fieldModel.getTitleStyleName(), null, createHelper);
		// merge column
		if (cacheModel.getExcelWrite().mergeTitleRowNum() > ExcelConstant.ZERO_SHORT || cacheModel.getExcelWrite().mergeTitleColNum() > ExcelConstant.ZERO_SHORT) {
			int mergeRowEndIndex = rowIndex + cacheModel.getExcelWrite().mergeTitleRowNum();
			int mergeColEndIndex = colIndex + cacheModel.getExcelWrite().mergeTitleColNum();

			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));

			if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
				colIndex = mergeColEndIndex;
			} else {
				rowIndex = mergeRowEndIndex;
			}

		}
		if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex++;
		} else {
			rowIndex++;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表标题
	 *
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param initColIndex
	 * @param initRowIndex
	 * @param fillFieldModelList
	 * @return
	 */
	protected ColPointDto fillListTitle(ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, final int initColIndex, final int initRowIndex, ExcelWriterModel exportModel, final List<ExcelCacheFieldModel> fillFieldModelList) {
		if (cacheModel.getExcelWrite().fillTitle()) {
			for (ExcelCacheFieldModel fieldModel : fillFieldModelList) {
				if (fieldModel.isMap()) {
					ColPointDto colPointDto = fillListMapTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, fieldModel, exportModel);
					rowIndex = colPointDto.rowIndex;
					colIndex = colPointDto.colIndex;
				} else {
					ColPointDto colPointDto = fillListBeanTitle(cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, fieldModel);
					rowIndex = colPointDto.rowIndex;
					colIndex = colPointDto.colIndex;
				}
			}
			if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
				rowIndex += cacheModel.getExcelWrite().mergeTitleRowNum() + ExcelConstant.ONE_INT;
			} else {
				colIndex += cacheModel.getExcelWrite().mergeTitleColNum() + ExcelConstant.ONE_INT;
			}
			if (cacheModel.getExcelWrite().freezeTitle()) {
				sheet.createFreezePane(ExcelConstant.ZERO_SHORT, cacheModel.getExcelWrite().rowIndex() + ExcelConstant.ONE_INT, ExcelConstant.ZERO_SHORT, cacheModel.getExcelWrite().rowIndex() + ExcelConstant.ONE_INT);
			}
			if (cacheModel.getExcelWrite().filterTitle()) {
				CellRangeAddress filterRange = new CellRangeAddress(initRowIndex, rowIndex, initColIndex, colIndex-ExcelConstant.ONE_INT);
				sheet.setAutoFilter(filterRange);
			}
		}

		if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex = initColIndex;
		} else {
			rowIndex = initRowIndex;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表list字段
	 *
	 * @param workbook
	 * @param cacheFieldModel
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param value
	 * @param styleName
	 * @return
	 */
	protected ColPointDto fillListField(Workbook workbook, ExcelCacheFieldModel cacheFieldModel, ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, Object value, String styleName, String linkName) {
		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setRowHeight(row, cacheFieldModel.getListRowHeight());
		ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
		if (Objects.nonNull(value) && value instanceof byte[]) {
			createPicture(workbook, sheet, cell, (byte[]) value, styleName);
		} else {
			setCellValueAndStyle(cell, value, styleName, linkName, createHelper);
		}
		// merge column   不判断标题是否合并
		if (canMergeColumn(cacheModel.getExcelWrite(), fillStyleEnum)) {
			int mergeContentRowNum = cacheModel.getExcelWrite().mergeContentRowNum(), mergeContentColNum = cacheModel.getExcelWrite().mergeContentColNum();

			if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
				mergeContentColNum = cacheModel.getExcelWrite().mergeTitleColNum();
			} else if (ExcelWriterFillStyleEnum.HORIZONTAL == fillStyleEnum) {
				mergeContentRowNum = cacheModel.getExcelWrite().mergeTitleRowNum();
			}
			int mergeRowEndIndex = rowIndex + mergeContentRowNum;
			int mergeColEndIndex = colIndex + mergeContentColNum;
			sheet.addMergedRegion(new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex));
			if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
				colIndex = mergeColEndIndex;
			} else {
				rowIndex = mergeRowEndIndex;
			}
		}
		if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
			colIndex++;
		} else {
			rowIndex++;
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 是否可以合并list列表单元格
	 *
	 * @param excelWrite
	 * @param fillStyleEnum
	 * @return
	 */
	protected boolean canMergeColumn(ExcelWrite excelWrite, ExcelWriterFillStyleEnum fillStyleEnum) {
		if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
			if (excelWrite.mergeTitleColNum() > ExcelConstant.ZERO_SHORT || excelWrite.mergeContentColNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			} else if (excelWrite.mergeContentRowNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			}
		} else if (ExcelWriterFillStyleEnum.HORIZONTAL == fillStyleEnum) {
			if (excelWrite.mergeTitleRowNum() > ExcelConstant.ZERO_SHORT || excelWrite.mergeContentRowNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			} else if (excelWrite.mergeContentColNum() > ExcelConstant.ZERO_SHORT) {
				return true;
			}
		}
		return false;
	}

	/**
	 * 填充列表map字段
	 *
	 * @param workbook
	 * @param exportCell
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param value
	 * @param styleName
	 * @param cacheFieldModel
	 * @return
	 */
	protected ColPointDto fillListMapField(Workbook workbook, ExcelWriteProperty exportCell, Sheet sheet, CreationHelper createHelper, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, Object value, String styleName, ExcelCacheFieldModel cacheFieldModel) {
		if (value instanceof Map) {
			Iterator keyIterator = ((Map) value).keySet().iterator();
			while (keyIterator.hasNext()) {
				Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setRowHeight(row, cacheFieldModel.getContentRowHeight());
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				Object cellValue = ((Map) value).get(keyIterator.next());
				if (Objects.nonNull(cellValue) && cellValue instanceof byte[]) {
					createPicture(workbook, sheet, cell, (byte[]) cellValue, styleName);
				} else {
					setCellValueAndStyle(cell, cellValue, styleName, exportCell.linkName(), createHelper);
				}
				if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
					colIndex++;
				} else {
					rowIndex++;
				}
			}
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充列表内容
	 *
	 * @param workbook
	 * @param exportModel
	 * @param cacheModel
	 * @param sheet
	 * @param createHelper
	 * @param rowIndex
	 * @param colIndex
	 * @param fillStyleEnum
	 * @param initColIndex
	 * @param initRowIndex
	 * @return
	 */
	protected ColPointDto fillListContent(Workbook workbook, ExcelWriterModel exportModel, ExcelCacheModel cacheModel, Sheet sheet, CreationHelper createHelper, List<ExcelCacheFieldModel> fillFieldModelList, int rowIndex, int colIndex, ExcelWriterFillStyleEnum fillStyleEnum, final int initColIndex, final int initRowIndex) {
		int contentRow = ExcelConstant.ZERO_SHORT;
		AtomicLong incrementSeq = new AtomicLong(ExcelConstant.ZERO_SHORT);
		// 填充内容
		for (ExcelBaseModel model : exportModel.getDataModelList()) {
			contentRow++;
			try {
				for (ExcelCacheFieldModel cacheFieldModel : fillFieldModelList) {
					Object value ;
					if (Objects.nonNull(cacheFieldModel.getGetMethod())) {
						value = cacheFieldModel.getGetMethod().invoke(model);
					}else{
						value = incrementSeq.incrementAndGet();
					}
					ExcelWriteProperty exportCell = cacheFieldModel.getExportCell();
					String linkName = resolveLinkName(cacheFieldModel, model);
					String styleName = contentRow % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();

					if(Objects.nonNull(exportCell)) {
						ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
						value = convertRuntimeValue(cacheFieldModel, value);
						value = formatValue(value, exportCell.formatPattern(), formatter);
						value = addValidatorOrComment(workbook, sheet, colIndex, rowIndex, exportCell, value);
					}
					if (cacheFieldModel.isMap()) {
						ColPointDto colPointDto = fillListMapField(workbook, exportCell, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, value, styleName, cacheFieldModel);
						rowIndex = colPointDto.rowIndex;
						colIndex = colPointDto.colIndex;
					} else {
						ColPointDto colPointDto = fillListField(workbook, cacheFieldModel, cacheModel, sheet, createHelper, rowIndex, colIndex, fillStyleEnum, value, styleName, linkName);
						rowIndex = colPointDto.rowIndex;
						colIndex = colPointDto.colIndex;
					}

				}
				if (ExcelWriterFillStyleEnum.VERTICAL == fillStyleEnum) {
					rowIndex += cacheModel.getExcelWrite().mergeContentRowNum() + ExcelConstant.ONE_INT;
					colIndex = initColIndex;
				} else {
					colIndex += cacheModel.getExcelWrite().mergeContentColNum() + ExcelConstant.ONE_INT;
					rowIndex = initRowIndex;
				}
			} catch (Exception e) {
				log.info(Throwables.getStackTraceAsString(e));
				throw new ExcelWriterException(e.getMessage());
			}
		}
		return ColPointDto.builder().colIndex(colIndex).rowIndex(rowIndex).build();
	}

	/**
	 * 填充内容
	 *
	 * @param wb              Workbook
	 * @param sheet           sheet 名称
	 * @param colIndex        列宽
	 * @param exportCell      cell
	 * @param value           value
	 * @param row             row
	 * @param createHelper    excelHelper
	 * @param excelWrite
	 * @param cacheFieldModel
	 * @return int
	 */
	protected ColPointDto fillContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelWriteProperty exportCell, Object value, Row row, CreationHelper createHelper, ExcelWrite excelWrite, boolean isMap, ExcelCacheFieldModel cacheFieldModel, String linkName) {

		ExcelWriterCellTitleModelEnum titleModelEnum = excelWrite.titleModel();
		if (titleModelEnum == ExcelWriterCellTitleModelEnum.DEFAULT) {
			titleModelEnum = exportCell.titleModel();
		} else {
			titleModelEnum = exportCell.titleModel() == ExcelWriterCellTitleModelEnum.DEFAULT ? titleModelEnum : exportCell.titleModel();
		}

		if (titleModelEnum == ExcelWriterCellTitleModelEnum.STAND_ALONE || titleModelEnum == ExcelWriterCellTitleModelEnum.DEFAULT) {
			return fillStandAloneContent(wb, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelWrite, isMap, cacheFieldModel, linkName);
		} else if (titleModelEnum == ExcelWriterCellTitleModelEnum.WITH_VALUE) {
			return fillWithValueContent(wb, sheet, colIndex, rowIndex, exportCell, value, row, createHelper, excelWrite, cacheFieldModel, linkName);
		}
		return null;
	}

	protected ColPointDto fillStandAloneContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelWriteProperty exportCell,
												Object value, Row row, CreationHelper createHelper, ExcelWrite excelWrite,
												boolean isMap, ExcelCacheFieldModel cacheFieldModel, String linkName) {
		if (isMap && value instanceof Map) {
			return fillListMapField(wb, exportCell, sheet, createHelper, rowIndex, colIndex, exportCell.fillStyle(),
					value, cacheFieldModel.getContentStyleName(), cacheFieldModel);
		}
		String titleName = resolveRuntimeTitle(cacheFieldModel);
		if (StringUtil.notEmpty(titleName)) {
			Cell titleCell = ExcelHelper.getCellOrCreate(row, colIndex);
			ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
			MergeTitleDto mergedTitle = cellMergedTitle(exportCell, rowIndex, colIndex, sheet, cacheFieldModel.getTitleStyleName(), wb);
			fillMergeCellValueAndStyle(titleCell, titleName + exportCell.separator(),
					cacheFieldModel.getTitleStyleName(), null, createHelper, mergedTitle);
			colIndex = mergedTitle.getPoint().getColIndex() + ExcelConstant.ONE_INT;
		}
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
		value = addValidatorOrComment(wb, sheet, colIndex, rowIndex, exportCell, value);
		if (value instanceof byte[]) {
			createPicture(wb, sheet, cell, (byte[]) value, cacheFieldModel.getContentStyleName());
		} else if (excelWrite.fillContent()) {
			setCellValueAndStyle(cell, value, cacheFieldModel.getContentStyleName(), linkName, createHelper);
		}
		return cellMerged(exportCell, rowIndex, colIndex, sheet, cacheFieldModel.getContentStyleName(), value);
	}

	protected ColPointDto fillWithValueContent(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelWriteProperty exportCell,
											   Object value, Row row, CreationHelper createHelper, ExcelWrite excelWrite,
											   ExcelCacheFieldModel cacheFieldModel, String linkName) {
		Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
		ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
		value = addValidatorOrComment(wb, sheet, colIndex, rowIndex, exportCell, value);
		if (value instanceof byte[]) {
			createPicture(wb, sheet, cell, (byte[]) value, cacheFieldModel.getTitleStyleName());
			return cellMerged(exportCell, rowIndex, colIndex, sheet, cacheFieldModel.getTitleStyleName(), value);
		}
		String titleName = resolveRuntimeTitle(cacheFieldModel);
		if (StringUtil.notEmpty(titleName)) {
			String content = titleName + exportCell.separator() + Optional.ofNullable(value).orElse(ExcelConstant.NULL_STR);
			int titleLen = titleName.length();
			ExcelRichTextModel titleModel = null;
			ExcelRichTextModel contentModel = null;
			CellStyle titleStyle = styleLocal.get().get(cacheFieldModel.getTitleStyleName());
			if (titleStyle != null) {
				titleModel = new ExcelRichTextModel();
				titleModel.setFont(wb.getFontAt(titleStyle.getFontIndex()));
				titleModel.setStartIndex(ExcelConstant.ZERO_SHORT);
				titleModel.setEndIndex(titleLen);
			}
			CellStyle contentStyle = styleLocal.get().get(cacheFieldModel.getContentStyleName());
			if (contentStyle != null) {
				contentModel = new ExcelRichTextModel();
				contentModel.setFont(wb.getFontAt(contentStyle.getFontIndex()));
				contentModel.setStartIndex(titleLen);
				contentModel.setEndIndex(content.length());
			}
			RichTextString richText = ExcelHelper.createRichText(createHelper, content, titleModel, contentModel);
			setCellValueAndStyle(cell, richText, cacheFieldModel.getTitleStyleName(), null, createHelper);
		} else {
			setCellValueAndStyle(cell, value, cacheFieldModel.getTitleStyleName(), linkName, createHelper);
		}
		return cellMerged(exportCell, rowIndex, colIndex, sheet, cacheFieldModel.getTitleStyleName(), value);
	}

	protected ColPointDto cellMerged(ExcelWriteProperty exportCell, int rowIndex, int colIndex, Sheet sheet, String styleName, Object value) {
		int mergeRowNum = exportCell.mergeRowNum();
		int mergeColNum = exportCell.mergeContentColNum();
		if (mergeRowNum > ExcelConstant.ZERO_SHORT || mergeColNum > ExcelConstant.ZERO_SHORT) {
			int lastRowIndex = rowIndex + mergeRowNum;
			int lastColIndex = colIndex + mergeColNum;
			CellRangeAddress cellAddresses = new CellRangeAddress(rowIndex, lastRowIndex, colIndex, lastColIndex);
			sheet.addMergedRegion(cellAddresses);
			CellStyle cellStyle = styleLocal.get().get(styleName);
			if (cellStyle != null) {
				RegionUtil.setBorderTop(cellStyle.getBorderTop(), cellAddresses, sheet);
				RegionUtil.setBorderBottom(cellStyle.getBorderBottom(), cellAddresses, sheet);
				RegionUtil.setBorderRight(cellStyle.getBorderRight(), cellAddresses, sheet);
				RegionUtil.setBorderLeft(cellStyle.getBorderLeft(), cellAddresses, sheet);
			}
			rowIndex = lastRowIndex;
			colIndex = lastColIndex;
		}
		return ColPointDto.builder().rowIndex(rowIndex).colIndex(colIndex).build();
	}

	protected void fillMergeCellValueAndStyle(Cell cell, Object value, String styleName, String linkName,
											  CreationHelper createHelper, MergeTitleDto mergeTitleDto) {
		if (mergeTitleDto.getRangeAddress().getLastRow() > ExcelConstant.ZERO_SHORT
				|| mergeTitleDto.getRangeAddress().getLastColumn() > ExcelConstant.ZERO_SHORT) {
			for (int currentRow = mergeTitleDto.getRangeAddress().getFirstRow(); currentRow <= mergeTitleDto.getRangeAddress().getLastRow(); currentRow++) {
				Row styleRow = ExcelHelper.getRowOrCreate(cell.getSheet(), currentRow);
				for (int currentCol = mergeTitleDto.getRangeAddress().getFirstColumn(); currentCol <= mergeTitleDto.getRangeAddress().getLastColumn(); currentCol++) {
					Cell styleCell = ExcelHelper.getCellOrCreate(styleRow, currentCol);
					setStyle(styleCell, value, styleName, linkName);
				}
			}
		} else {
			setStyle(cell, value, styleName, linkName);
		}
		Object cellValue = ExcelHelper.createHyperlink(cell, Objects.isNull(value) ? ExcelConstant.NULL_STR : value, linkName, createHelper);
		ExcelHelper.setCellValue(cell, cellValue);
	}

	protected Object addValidatorOrComment(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelWriteProperty exportCell, Object value) {
		List<String> dropDownOptions = Arrays.asList(exportCell.dropDownOptions());
		String validationTitle = null;
		String validationMessage = null;
		if (value instanceof ExcelWriterComboParam) {
			ExcelWriterComboParam comboParam = (ExcelWriterComboParam) value;
			dropDownOptions = comboParam.getOptions();
			validationTitle = comboParam.getTitle();
			validationMessage = comboParam.getMessage();
			value = comboParam.getValue();
		}
		if (!dropDownOptions.isEmpty()) {
			ExcelHelper.addDropDownValidation(sheet, new CellRangeAddressList(rowIndex, rowIndex, colIndex, colIndex), dropDownOptions, validationTitle, validationMessage);
		}
		if (value instanceof ExcelWriterNumberScopeParam) {
			ExcelWriterNumberScopeParam scopeParam = (ExcelWriterNumberScopeParam) value;
			ExcelHelper.addRangeValidation(sheet, new CellRangeAddressList(rowIndex, rowIndex, colIndex, colIndex),
					scopeParam.getStart(), scopeParam.getEnd(), scopeParam.getTitle(), scopeParam.getMessage());
			value = scopeParam.getValue();
		}
		return createComment(wb, sheet, colIndex, rowIndex, exportCell, value);
	}

	protected Object createComment(Workbook wb, Sheet sheet, int colIndex, int rowIndex, ExcelWriteProperty exportCell, Object value) {
		String commentText = exportCell.commentText();
		String commentFontName = exportCell.commentFontName();
		String author = ExcelConstant.NULL_STR;
		if (value instanceof ExcelWriterCommentParam) {
			ExcelWriterCommentParam commentParam = (ExcelWriterCommentParam) value;
			commentText = StringUtil.notEmpty(commentParam.getCommentText()) ? commentParam.getCommentText() : commentText;
			commentFontName = StringUtil.notEmpty(commentParam.getCommentFontName()) ? commentParam.getCommentFontName() : commentFontName;
			author = commentParam.getAuthor();
			value = commentParam.getValue();
		}
		if (StringUtil.notEmpty(commentText)) {
			ExcelHelper.createComment(wb, sheet, rowIndex, colIndex, author, commentText, fontLocal.get().get(commentFontName));
		}
		return value;
	}



	protected String resolveLinkName(ExcelCacheFieldModel cacheFieldModel, Object model) {
		ExcelWriteProperty exportCell = cacheFieldModel.getExportCell();
		if (Objects.isNull(exportCell)) {
			return ExcelConstant.NULL_STR;
		}
		Method linkNameGetMethod = cacheFieldModel.getLinkNameGetMethod();
		if (Objects.isNull(linkNameGetMethod) || Objects.isNull(model)) {
			return exportCell.linkName();
		}
		try {
			Object linkName = linkNameGetMethod.invoke(model);
			return Objects.isNull(linkName) ? exportCell.linkName() : linkName.toString();
		} catch (IllegalAccessException | InvocationTargetException e) {
			log.error("resolve hyperlink display name failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		}
	}

	protected List<ExcelCacheFieldModel> filterListFieldModels(ExcelCacheModel cacheModel, String[] excludeFields) {
		Map<String, String> excludeFieldMap = new HashMap<>();
		if (Objects.nonNull(excludeFields)) {
			for (String excludeField : excludeFields) {
				excludeFieldMap.put(excludeField, ExcelConstant.NULL_STR);
			}
		}
		List<ExcelCacheFieldModel> cacheFieldModelList = new ArrayList<>();
		for (ExcelCacheFieldModel fieldModel : cacheModel.getFieldModelList()) {
			if (Objects.nonNull(excludeFieldMap.get(fieldModel.getFieldName()))) {
				continue;
			}
			if (skipRuntimeField(fieldModel)) {
				continue;
			}
			cacheFieldModelList.add(fieldModel);
		}
		return cacheFieldModelList;
	}

	protected boolean skipRuntimeField(ExcelCacheFieldModel fieldModel) {
		String fieldName = fieldModel.getFieldName();
		if (runtimeIncludeFields != null && !runtimeIncludeFields.isEmpty() && !runtimeIncludeFields.contains(fieldName)) {
			return true;
		}
		return runtimeOnlyAlias && (runtimeHeaderAliases == null || !runtimeHeaderAliases.containsKey(fieldName));
	}

	protected String resolveRuntimeTitle(ExcelCacheFieldModel fieldModel) {
		if (runtimeHeaderAliases == null) {
			return fieldModel.getTitleName();
		}
		return runtimeHeaderAliases.getOrDefault(fieldModel.getFieldName(), fieldModel.getTitleName());
	}

	protected Object convertRuntimeValue(ExcelCacheFieldModel fieldModel, Object value) {
		if (runtimeFieldConverters != null) {
			java.util.function.Function<Object, Object> fieldConverter = runtimeFieldConverters.get(fieldModel.getFieldName());
			if (fieldConverter != null) {
				return fieldConverter.apply(value);
			}
		}
		if (value == null || runtimeTypeConverters == null) {
			return value;
		}
		for (Map.Entry<Class<?>, java.util.function.Function<Object, Object>> entry : runtimeTypeConverters.entrySet()) {
			if (entry.getKey().isAssignableFrom(value.getClass())) {
				return entry.getValue().apply(value);
			}
		}
		return value;
	}

	protected void writeFlatListTitleRow(Sheet sheet, CreationHelper creationHelper, List<ExcelCacheFieldModel> cacheFieldModelList, int rowIndex) {
		Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
		int colIndex = ExcelConstant.ZERO_SHORT;
		for (ExcelCacheFieldModel fieldModel : cacheFieldModelList) {
			Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
			ExcelHelper.setColWidth(sheet, colIndex, fieldModel.getColWidth());
			ExcelHelper.setRowHeight(row, fieldModel.getTitleRowHeight());
			setCellValueAndStyle(cell, resolveRuntimeTitle(fieldModel), fieldModel.getTitleStyleName(), null, creationHelper);
			++colIndex;
		}
	}

	protected void writeFlatListContentRow(Workbook workbook, Sheet sheet, CreationHelper creationHelper,
										   List<ExcelCacheFieldModel> cacheFieldModelList, ExcelBaseModel model,
										   AtomicLong incrementSeq, int rowIndex, int contentRowNumber) {
		int colIndex = ExcelConstant.ZERO_SHORT;
		for (ExcelCacheFieldModel cacheFieldModel : cacheFieldModelList) {
			Object value = resolveFlatListCellValue(cacheFieldModel, model, incrementSeq);
			String linkName = resolveLinkName(cacheFieldModel, model);
			String styleName = contentRowNumber % ExcelConstant.TOW_INT == ExcelConstant.ZERO_SHORT ? cacheFieldModel.getEvenRowStyleName() : cacheFieldModel.getContentStyleName();
			Row row = ExcelHelper.getRowOrCreate(sheet, rowIndex);
			Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
			ExcelHelper.setRowHeight(row, cacheFieldModel.getListRowHeight());
			ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
			if (Objects.nonNull(value) && value instanceof byte[]) {
				createPicture(workbook, sheet, cell, (byte[]) value, styleName);
			} else {
				setCellValueAndStyle(cell, value, styleName, linkName, creationHelper);
			}
			colIndex++;
		}
	}

	protected Object resolveFlatListCellValue(ExcelCacheFieldModel cacheFieldModel, ExcelBaseModel model, AtomicLong incrementSeq) {
		ExcelWriteProperty exportCell = cacheFieldModel.getExportCell();
		if (Objects.isNull(exportCell)) {
			return incrementSeq.incrementAndGet();
		}
		try {
			Object value = cacheFieldModel.getGetMethod().invoke(model);
			ExcelWriterDataFormat formatter = getFormatter(exportCell.formatter());
			value = convertRuntimeValue(cacheFieldModel, value);
			return formatValue(value, exportCell.formatPattern(), formatter);
		} catch (IllegalAccessException | InvocationTargetException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		}
	}

	/**
	 * 填充合并单元格
	 *
	 * @param exportCell
	 * @param rowIndex
	 * @param colIndex
	 * @param sheet
	 * @param titleStyleName
	 * @return
	 */
	protected MergeTitleDto cellMergedTitle(ExcelWriteProperty exportCell, int rowIndex, int colIndex, Sheet sheet, String titleStyleName , Workbook wb) {
		int mergeContentRowNum = exportCell.mergeRowNum(), mergeContentColNum = exportCell.mergeTitleColNum();
		int mergeRowEndIndex = rowIndex + mergeContentRowNum;
		int mergeColEndIndex = colIndex + mergeContentColNum;
		CellRangeAddress cellAddresses = new CellRangeAddress(rowIndex, mergeRowEndIndex, colIndex, mergeColEndIndex);
		if (mergeContentRowNum > ExcelConstant.ZERO_SHORT || mergeContentColNum > ExcelConstant.ZERO_SHORT) {

			CellStyle cellStyle = styleLocal.get().get(titleStyleName);
			if(Objects.nonNull(cellStyle)) {
				if(sheet instanceof HSSFSheet) {
					RegionUtil.setBorderTop(cellStyle.getBorderTop(), cellAddresses, sheet);
					RegionUtil.setBorderBottom(cellStyle.getBorderBottom(), cellAddresses, sheet);
					RegionUtil.setBorderRight(cellStyle.getBorderRight(), cellAddresses, sheet);
					RegionUtil.setBorderLeft(cellStyle.getBorderLeft(), cellAddresses, sheet);
					RegionUtil.setTopBorderColor(cellStyle.getTopBorderColor(), cellAddresses, sheet);
					RegionUtil.setBottomBorderColor(cellStyle.getBottomBorderColor(), cellAddresses, sheet);
					RegionUtil.setLeftBorderColor(cellStyle.getLeftBorderColor(), cellAddresses, sheet);
					RegionUtil.setRightBorderColor(cellStyle.getRightBorderColor(), cellAddresses, sheet);
				}
			}
			sheet.addMergedRegion(cellAddresses);
			colIndex = mergeColEndIndex;
		}
		MergeTitleDto titleDto = new MergeTitleDto();
		titleDto.setPoint(ColPointDto.builder().rowIndex(rowIndex).colIndex(colIndex).build());
		titleDto.setRangeAddress(cellAddresses);
		return titleDto;
	}




	protected Object setStyle(Cell cell, Object value, String styleName, String linkName) {
		CellStyle cellStyle = null;
		if (StringUtil.notEmpty(styleName)) {
			cellStyle = styleLocal.get().get(styleName);
			if (cellStyle != null) {
				cell.setCellStyle(cellStyle);
			}
		}
		if (value == null) {
			return null;
		}
		if (cellStyle == null) {
			if (value instanceof Date || value instanceof Calendar) {
				cell.setCellStyle(styleLocal.get().get(ExcelBasicStyle.STYLE_DATE_YYYYMMDDHHMMSS));
			} else if (StringUtil.notEmpty(linkName)) {
				cell.setCellStyle(styleLocal.get().get(ExcelBasicStyle.STYLE_HLINK));
			}
		}
		return value;
	}

	/**
	 * 设置单元格内容和样式
	 *
	 * @param cell         单元格
	 * @param value        数据
	 * @param styleName    样式名称
	 * @param linkName     连接名称
	 * @param createHelper excelHelper
	 * @param expression   表达式
	 */
	protected void setCellValueAndStyle(Cell cell, Object value, String styleName, String linkName, CreationHelper createHelper, String expression) {
		value = setStyle(cell, value, styleName, linkName);
		value = Objects.isNull(value) ? ExcelConstant.NULL_STR : value;
		value = ExcelHelper.createHyperlink(cell, value, linkName, createHelper);
		replaceExpression(cell, value, expression);
	}
	/**
	 * 设置单元格内容和样式（不替换表达式）
	 *
	 * @param cell         单元格
	 * @param value        数据
	 * @param styleName    样式名称
	 * @param linkName     连接名称
	 * @param createHelper excelHelper
	 */
	protected void setCellValueAndStyle(Cell cell, Object value, String styleName, String linkName, CreationHelper createHelper) {
		value = setStyle(cell, value, styleName, linkName);
		value = Objects.isNull(value) ? ExcelConstant.NULL_STR : value;
		value = ExcelHelper.createHyperlink(cell, value, linkName, createHelper);
		ExcelHelper.setCellValue(cell, value);
	}

	/**
	 * 替换标签
	 *
	 * @param cell
	 * @param value
	 * @param expression
	 */
	protected void replaceExpression(Cell cell, Object value, String expression) {
		value = Objects.isNull(value) ? ExcelConstant.NULL_STR : value;
		String expressionValue = cell.getStringCellValue();
		if (expressionValue.matches(expression)) {
			ExcelHelper.setCellValue(cell, value);
		} else {
			String cellValue = expressionValue.replaceAll(expression, value.toString());
			cell.setCellValue(cellValue);
		}
	}





	/**
	 * 格式化数据
	 *
	 * @param data          数据
	 * @param formatPattern 格式化字符串
	 * @param formatter     格式化器
	 * @return Object
	 */
	protected Object formatValue(Object data, String formatPattern, ExcelWriterDataFormat formatter) {
		if (null == data || StringUtil.isEmpty(formatPattern)) {
			return data;
		}
		if (null != formatter) {
			return formatter.format(data, formatPattern);
		}
		return dataFormat.format(data, formatPattern);
	}


	/**
	 * 创建图片
	 *
	 * @param wb    WorkBook
	 * @param sheet sheet
	 * @param cell  单元格
	 * @param bytes 图片流数组
	 * @return Picture
	 */
	protected Picture createPicture(Workbook wb, Sheet sheet, Cell cell, byte[] bytes, String styleName) {
		// 设置样式
		if (StringUtil.notEmpty(styleName)) {
			CellStyle cellStyle = styleLocal.get().get(styleName);
			if (null != cellStyle) {
				cell.setCellStyle(cellStyle);
			}
		}
		int i = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		// 创建锚点
		CreationHelper creationHelper = wb.getCreationHelper();
		ClientAnchor clientAnchor = creationHelper.createClientAnchor();
		clientAnchor.setRow1(cell.getRowIndex());
		clientAnchor.setCol1(cell.getColumnIndex());
		clientAnchor.setRow2(cell.getRowIndex() + ExcelConstant.ONE_INT);
		clientAnchor.setCol2(cell.getColumnIndex() + ExcelConstant.ONE_INT);

		Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();
		return drawingPatriarch.createPicture(clientAnchor, i);
	}


	/**
	 * 添加空数据默认内容
	 *
	 * @param wb             workbook
	 * @param creationHelper creationHelper
	 */
	protected void addNoResultData(Workbook wb, CreationHelper creationHelper) {
		int numberOfSheets = wb.getNumberOfSheets();
		// sheet数量大于1说明有数据
		if (numberOfSheets > ExcelConstant.ONE_INT) {
			return ;
		}
		// 如果等于0则新创建sheet，否则获取第一个sheet
		Sheet sheet = wb.getNumberOfSheets() == ExcelConstant.ZERO_SHORT ? wb.createSheet(ExcelConstant.DEFAULT_SHEET_NAME) : wb.getSheetAt(ExcelConstant.ZERO_SHORT);
		if (sheet.getLastRowNum() > ExcelConstant.ZERO_SHORT || sheet.getPhysicalNumberOfRows() > ExcelConstant.ZERO_SHORT) {
			return;
		}
		if (Objects.isNull(listCla)) {
			// 判断sheetName
			ExcelWriterCellParam columnModel = new ExcelWriterCellParam();
			columnModel.setRowIndex(0);
			columnModel.setColIndex(0);
			columnModel.setSheetName(sheet.getSheetName());
			if(noneDataTips) {
				columnModel.setColWidth(ExcelConstant.SHORT_500);
				columnModel.setRowHeight(ExcelConstant.SHORT_50);
				columnModel.setStyleName(ExcelBasicStyle.STYLE_ZEBRA_TITLE_ROW);
				columnModel.setValue(ExcelErrorMsgConstant.ERROR_EXPORT_NOT_FOUND_DATA);
			}
			fillColumn(wb, creationHelper, columnModel, columnModel.getRowIndex(), columnModel.getColIndex());
		} else {
			ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(listCla);
			Row row = ExcelHelper.getRowOrCreate(sheet, ExcelConstant.ZERO_SHORT);
			int colIndex = ExcelConstant.ZERO_SHORT;
			for (ExcelCacheFieldModel cacheFieldModel : cacheModel.getFieldModelList()) {
				if (skipRuntimeField(cacheFieldModel)) {
					continue;
				}
				Object titleName = resolveRuntimeTitle(cacheFieldModel);
				if (Objects.isNull(titleName)) {
					continue;
				}
				Cell cell = ExcelHelper.getCellOrCreate(row, colIndex);
				ExcelHelper.setColWidth(sheet, colIndex, cacheFieldModel.getColWidth());
				ExcelHelper.setRowHeight(row, cacheFieldModel.getTitleRowHeight());

				this.setCellValueAndStyle(cell, titleName, cacheFieldModel.getTitleStyleName(), null, creationHelper);
				colIndex++;
			}
		}
	}

	@lombok.Data
	@lombok.Builder
	protected static class ColPointDto {
		private int rowIndex;
		private int colIndex;
	}

	@lombok.Data
	protected static class MergeTitleDto {
		private ColPointDto point;
		private CellRangeAddress rangeAddress;
	}

	protected void clearData(){
		exportModelList = new ArrayList<>();
		exportBeanList = new ArrayList<>();
		customColumnModelList = new ArrayList<>();
		mergeCustomColumnModelList = new ArrayList<>();
		styleList = new ArrayList<>(ExcelConstant.TOW_INT);
		excludeFieldMap = new HashMap<>();
		commentMap = new HashMap<>();
		conditionalStyleMap = new HashMap<>();
		conditionalStyleList = new ArrayList<>();
		runtimeHeaderAliases = new LinkedHashMap<>();
		runtimeIncludeFields = new LinkedHashSet<>();
		runtimeOnlyAlias = false;
		runtimeTypeConverters = new LinkedHashMap<>();
		runtimeFieldConverters = new LinkedHashMap<>();
		dataFormatMap = new HashMap<>();
		customWrite = null;
		listCla = null;
		incrementSeqMap = Maps.newConcurrentMap();
		fontLocal = ThreadLocal.withInitial(Maps::newHashMap);
		styleLocal = ThreadLocal.withInitial(Maps::newHashMap);
		colorLocal = ThreadLocal.withInitial(Maps::newHashMap);
	}
}
