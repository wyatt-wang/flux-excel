package com.github.excel.write;

import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.helper.ExcelHelper;
import com.github.excel.write.pipeline.ExcelWriteContext;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.Objects;

public class ExcelWriteKernel {

	private final BaseExcelWriter writer;

	public ExcelWriteKernel() {
		this.writer = new BaseExcelWriter();
	}

	public void createWorkbook(ExcelWriteContext context) {
		writer.exportModelList = context.getExportModelList();
		writer.exportBeanList = context.getExportBeanList();
		writer.customColumnModelList = context.getCustomColumnModelList();
		writer.mergeCustomColumnModelList = context.getMergeCustomColumnModelList();
		writer.styleList = context.getStyleList();
		writer.excludeFieldMap = context.getExcludeFieldMap();
		writer.commentMap = context.getCommentMap();
		writer.conditionalStyleMap = context.getConditionalStyleMap();
		writer.conditionalStyleList = context.getConditionalStyleList();
		writer.runtimeHeaderAliases = context.getRuntimeHeaderAliases();
		writer.runtimeIncludeFields = context.getRuntimeIncludeFields();
		writer.runtimeOnlyAlias = context.isRuntimeOnlyAlias();
		writer.runtimeTypeConverters = context.getRuntimeTypeConverters();
		writer.runtimeFieldConverters = context.getRuntimeFieldConverters();
		writer.template = null;
		writer.customWrite = context.getCustomWrite();
		writer.streaming = context.isStreaming();
		writer.selectSheet = context.getSelectSheet();
		writer.listCla = context.getListCla();
		writer.noneDataTips = context.isNoneDataTips();
		writer.sheetRowMaxCount = context.getSheetRowMaxCount();
		Workbook workbook = com.github.excel.helper.WorkbookHelper.createWriteWorkBook(context.getWriterParam());
		ThreadLocal<WorkbookCachePool.WorkbookCacheModel> workbookThreadLocal = WorkbookCachePool.addBasicStyle(workbook);
		if (Objects.isNull(workbookThreadLocal)) {
			throw new ExcelWriterException("Failed to fetch workbook from cache");
		}
		WorkbookCachePool.WorkbookCacheModel cacheModel = workbookThreadLocal.get();
		writer.styleLocal.set(cacheModel.getStyleMap());
		writer.fontLocal.set(cacheModel.getFontMap());
		writer.colorLocal.set(cacheModel.getColorMap());
		Workbook cacheWorkbook = cacheModel.getWorkbook();
		context.setWorkbook(cacheWorkbook);
		if (cacheWorkbook instanceof SXSSFWorkbook) {
			context.setSxssfWorkbook((SXSSFWorkbook) cacheWorkbook);
		}
		context.setCreationHelper(cacheWorkbook.getCreationHelper());
	}

	public void prepareStyles(ExcelWriteContext context) {
		writer.initStyle(context.getWorkbook());
	}

	public void writeCustomCells(ExcelWriteContext context) {
		writer.fillCustomColumn(context.getWorkbook(), context.getCreationHelper());
		writer.fillMergeCustomColumn(context.getWorkbook(), context.getCreationHelper());
	}

	public void writeBeans(ExcelWriteContext context) throws IllegalAccessException, InvocationTargetException {
		for (var exportModel : context.getExportBeanList()) {
			var sheet = ExcelHelper.getSheetOrCreate(context.getWorkbook(), exportModel.getSheetName());
			writer.fillBean(context.getWorkbook(), context.getCreationHelper(), exportModel, sheet);
		}
	}

	public void writeLists(ExcelWriteContext context) {
		for (var exportModel : context.getExportModelList()) {
			var sheet = ExcelHelper.getSheetOrCreate(context.getWorkbook(), exportModel.getSheetName());
			writer.fillBeanList(context.getWorkbook(), context.getCreationHelper(), exportModel, sheet);
		}
	}

	public void applyPostProcessors(ExcelWriteContext context) {
		writer.applyConditionalStyles(context.getWorkbook());
		if (context.getCustomWrite() != null) {
			context.getCustomWrite().execute(context.getWorkbook());
		}
		writer.addNoResultData(context.getWorkbook(), context.getCreationHelper());
		writer.selectSheet(context.getWorkbook());
	}

	public void writeOutput(ExcelWriteContext context) throws IOException {
		context.getWorkbook().write(context.getOutputStream());
	}

	public void cleanup(ExcelWriteContext context) {
		try {
			if (Objects.nonNull(context.getWorkbook())) {
				context.getWorkbook().close();
			}
		} catch (IOException e) {
			throw new ExcelWriterException(e.getMessage());
		} finally {
			if (Objects.nonNull(context.getSxssfWorkbook())) {
				context.getSxssfWorkbook().dispose();
			}
			writer.clearData();
		}
	}
}
