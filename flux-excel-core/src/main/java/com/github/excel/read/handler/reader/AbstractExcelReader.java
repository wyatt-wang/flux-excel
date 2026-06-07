package com.github.excel.read.handler.reader;

import com.github.excel.boot.ExcelBootLoader;
import com.github.excel.constant.ExcelConstant;
import com.github.excel.context.ExcelReaderContext;
import com.github.excel.context.ExcelReaderModelContext;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.model.ExcelBaseModel;
import com.github.excel.model.ExcelCacheImportModel;
import com.github.excel.param.*;
import com.github.excel.read.facade.ExcelCustomReader;
import com.github.excel.read.handler.parser.ExcelReaderUserParser;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import lombok.extern.slf4j.Slf4j;

import java.util.List;
import java.util.ArrayList;
import java.util.Objects;
import java.util.Optional;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 抽象通用服务
 */
@Slf4j
public abstract class AbstractExcelReader implements ExcelReader {

	protected ExcelReaderContext readerContext = new ExcelReaderContext();
	protected Boolean failFast = Boolean.TRUE;
	protected String errorMsg = ExcelConstant.NULL_STR;

	public AbstractExcelReader(ExcelReaderStreamParam param) {
		this.readerContext.setReaderParam(param);

	}

	public AbstractExcelReader(ExcelReaderFileParam param) {
		String excelName = param.getFile().getName();
		if (StringUtil.isEmpty(excelName) || (!excelName.endsWith(ExcelConstant.XLSX_STR) && !excelName.endsWith(ExcelConstant.XLS_STR))) {
			throw new ExcelReaderException("excel suffix support only 'xls' and 'xlsx'");
		}
		this.readerContext.setReaderParam(param);

	}


	/**
	 * 创建model，用于addModel ｜ addModelList方法生成缓存信息
	 * @param baseParam 参数
	 * @param newInstance 是否创建实例，list 不用创建实例
	 * @return
	 */
	protected <T extends ExcelBaseModel> ExcelReaderModelContext<T> createModel(ExcelReaderDataParam<T> baseParam, boolean newInstance) {
		ExcelCacheImportModel excelCacheImportModel = ExcelBootLoader.getExcelCacheImportMapValue(baseParam.getModelCla());
		if (Objects.isNull(excelCacheImportModel)) {
			throw new ExcelReaderException("Model class unloaded");
		}
		ExcelReaderModelContext<T> modelContext = new ExcelReaderModelContext<>();
		try {
			if(newInstance) {
				modelContext.setModel(baseParam.getModelCla().newInstance());
			} else {
				modelContext.setModelList(new ArrayList<>());
			}
		} catch (InstantiationException | IllegalAccessException e) {
			log.error("Add model failed , cause :", Throwables.getStackTraceAsString(e));
			throw new ExcelReaderException("Add model failed , cause :" + e.getMessage());
		}
		modelContext.setModelCla(baseParam.getModelCla());
		modelContext.setCacheImportModel(excelCacheImportModel);
		modelContext.setParam(baseParam);
		return modelContext;
	}

	@Override
	public <T extends ExcelBaseModel> ExcelReader addModel(ExcelReaderModelParam<T> param) {
		ExcelReaderModelContext<T> modelContext = createModel(param, true);
		readerContext.getModelMap().put(param.getModelCla(), modelContext);
		return this;
	}

	@Override
	public <T extends ExcelBaseModel> ExcelReader addList(ExcelReaderListParam<T> param) {
		ExcelReaderModelContext<T> modelContext = createModel(param, false);
		readerContext.getModelMap().put(param.getModelCla(), modelContext);
		return this;
	}


	@Override
	public <T extends ExcelBaseModel> T getModel(Class<T> modelCla) {
		ExcelReaderModelContext<T> modelDto = (ExcelReaderModelContext<T>)readerContext.getModelMap().get(modelCla);
		return (T)Optional.ofNullable(modelDto).map(ExcelReaderModelContext::getModel).orElse(null);
	}

	@Override
	public <T extends ExcelBaseModel> List<T> getList(Class<T> modelCla) {
		ExcelReaderModelContext<T> modelDto = (ExcelReaderModelContext<T>)readerContext.getModelMap().get(modelCla);
		return Optional.ofNullable(modelDto).map(ExcelReaderModelContext::getModelList).orElse(null);
	}


	@Override
	public ExcelReader parse() {
		ExcelReaderUserParser handler = this.createHandler();
		handler.process();
		return this;
	}

	@Override
	public ExcelReader readCustom(ExcelCustomReader customReader) {
		readerContext.setExcelCustomReader(customReader);
		return this;
	}

	@Override
	public <T extends ExcelBaseModel> T parseAndGetModel(Class<T> modelCla) {
		parse();
		return getModel(modelCla);
	}

	@Override
	public <T extends ExcelBaseModel> List<T> parseAndGetList(Class<T> modelCla) {
		parse();
		return getList(modelCla);
	}

	@Override
	public ExcelReader setFailFast(Boolean failFast) {
		this.failFast = failFast;
		return this;
	}

	@Override
	public String getErrorMsg() {
		return errorMsg;
	}

	/**
	 * 创建处理器
	 * @return
	 */
	public abstract ExcelReaderUserParser createHandler();

}
