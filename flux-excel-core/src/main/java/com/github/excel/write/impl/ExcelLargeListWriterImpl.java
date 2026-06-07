package com.github.excel.write.impl;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.boot.WorkbookCachePool;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.helper.WorkbookHelper;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheModel;
import com.github.excel.model.ExcelCacheFieldModel;
import com.github.excel.write.style.AbstractExcelStyle;
import com.github.excel.write.BaseExcelWriter;
import com.github.excel.write.ExcelLargeListWriter;
import com.google.common.base.Throwables;
import com.google.common.collect.Maps;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: excel 导出大list数据实现
 */
@Slf4j
public class ExcelLargeListWriterImpl extends BaseExcelWriter implements ExcelLargeListWriter {

	private final SXSSFWorkbook workbook;
	private String sheetName;

	private boolean fillTitle = true;

	private int sheetNo = ExcelConstant.ONE_INT;

	private int rowIndex = ExcelConstant.ZERO_SHORT;


	public ExcelLargeListWriterImpl(String sheetName) {
		this(sheetName, ExcelConstant.INT_1000000, null, ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, false, false);
	}

	public ExcelLargeListWriterImpl(String sheetName, int rowAccessWindowSize,
									boolean compressTempFiles, boolean useSharedStringsTable) {
		this(sheetName, ExcelConstant.INT_1000000, null, rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
	}

	public ExcelLargeListWriterImpl(String sheetName, int sheetRowMaxCount) {
		this(sheetName, sheetRowMaxCount, null, ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, false, false);
	}

	public ExcelLargeListWriterImpl(String sheetName, int sheetRowMaxCount,
									int rowAccessWindowSize, boolean compressTempFiles, boolean useSharedStringsTable) {
		this(sheetName, sheetRowMaxCount, null, rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
	}

	public ExcelLargeListWriterImpl(String sheetName, int sheetRowMaxCount, Class<? extends ExcelBaseModel> listCla) {
		this(sheetName, sheetRowMaxCount, listCla, ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE, false, false);
	}

	public ExcelLargeListWriterImpl(String sheetName, int sheetRowMaxCount, Class<? extends ExcelBaseModel> listCla,
									int rowAccessWindowSize, boolean compressTempFiles, boolean useSharedStringsTable) {
		ThreadLocal<WorkbookCachePool.WorkbookCacheModel> sxssfWorkbookThreadLocal = WorkbookCachePool.addBasicStyle(
				WorkbookHelper.createStreamingXlsxWorkBook(rowAccessWindowSize, compressTempFiles, useSharedStringsTable)
		);
		WorkbookCachePool.WorkbookCacheModel workbookCacheModel = sxssfWorkbookThreadLocal.get();
		workbook = (SXSSFWorkbook) workbookCacheModel.getWorkbook();
		this.styleLocal.set(workbookCacheModel.getStyleMap());
		this.fontLocal.set(workbookCacheModel.getFontMap());
		if (sheetRowMaxCount > ExcelConstant.ZERO_SHORT && sheetRowMaxCount <= ExcelConstant.INT_1000000) {
			this.sheetRowMaxCount = sheetRowMaxCount;
		}else{
			this.sheetRowMaxCount = ExcelConstant.INT_1000000 ;
		}
		this.sheetName = sheetName;
		this.listCla = listCla;
	}

	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList,String[] excludeFields, Class<? extends ExcelBaseModel> modelClass){

		if (Objects.isNull(modelList) || modelList.isEmpty()) {
			return ;
		}

		ExcelCacheModel cacheModel = ExcelBootLoader.getExcelCacheMapValue(modelClass);
		CreationHelper creationHelper = workbook.getCreationHelper();
		SXSSFSheet sheet = null;
		List<ExcelCacheFieldModel> cacheFieldModelList = filterListFieldModels(cacheModel, excludeFields);

		// 填充内容
		int i = ExcelConstant.ZERO_SHORT;
		AtomicLong incrementSeq = incrementSeqMap.computeIfAbsent(modelClass, key -> new AtomicLong(ExcelConstant.ZERO_SHORT));
		for (T model : modelList) {
			// 填充标题
			if (fillTitle) {
				sheet = createNewSheet(creationHelper, cacheFieldModelList);
			} else {
				if(Objects.isNull(sheet)) {
					sheet = workbook.getSheet(sheetName + ExcelConstant.SHORT_TERM + sheetNo);
				}
			}
			i++;
			writeFlatListContentRow(workbook, sheet, creationHelper, cacheFieldModelList, model, incrementSeq, rowIndex, i);
			maxRowCreateSheet();
			rowIndex++;
		}

	}

	private SXSSFSheet createNewSheet(CreationHelper creationHelper, List<ExcelCacheFieldModel> cacheFieldModelList) {
		SXSSFSheet sheet;
		rowIndex = ExcelConstant.ZERO_SHORT;
		sheet = workbook.createSheet(sheetName + ExcelConstant.SHORT_TERM + sheetNo);
		writeFlatListTitleRow(sheet, creationHelper, cacheFieldModelList, rowIndex);
		fillTitle = false;
		++rowIndex;
		return sheet;
	}


	@Override
	public <T extends ExcelBaseModel> void process(List<T> modelList, Class<? extends ExcelBaseModel> modelCla){
		process(modelList, null, modelCla);
	}

	@Override
	public void setNoneDataTips(boolean noneDataTips) {
		this.noneDataTips = noneDataTips;
	}

	@Override
	public void export(OutputStream outputStream) {
		try {
			CreationHelper creationHelper = workbook.getCreationHelper();
			addNoResultData(workbook, creationHelper);
			workbook.write(outputStream);
		} catch (IOException e) {
			log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
			throw new ExcelWriterException(e.getMessage());
		}finally {
			try {
				outputStream.flush();
				outputStream.close();
			} catch (IOException e) {
				log.error("export failed cause:{}", Throwables.getStackTraceAsString(e));
				throw new ExcelWriterException(e.getMessage());
			}
			workbook.dispose();
			fontLocal.remove();
			styleLocal.remove();
			incrementSeqMap = Maps.newConcurrentMap();
		}
	}

	@Override
	public void export(File file, String fileName) {
		try (OutputStream outputStream = new FileOutputStream(file)) {
			export(outputStream);
		} catch (IOException e) {
			throw new ExcelWriterException(e);
		}
	}

	@Override
	public void addStyle(Class<? extends AbstractExcelStyle> styleClass) {
		initStyle(workbook,styleClass);
	}

	@Override
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	@Override
	public void setCla(Class<? extends ExcelBaseModel> cla) {
		this.listCla = cla;
	}

	@Override
	public void close() {

	}
	private void maxRowCreateSheet() {
		// 到达sheetRowMaxCount创建新的sheet
		if (rowIndex >= sheetRowMaxCount) {
			sheetNo++;
			fillTitle = true;
		}
	}
}
