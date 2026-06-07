package com.github.excel.read.handler.reader;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.param.ExcelReaderListParam;
import com.github.excel.param.ExcelReaderModelParam;
import com.github.excel.read.facade.ExcelCustomReader;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 读取
 */
public interface ExcelReader {
	/**
	 * 添加单个bean 解析
	 * @param param 参数
	 * @return
	 */
	<T extends ExcelBaseModel> ExcelReader addModel(ExcelReaderModelParam<T> param);

	default <T extends ExcelBaseModel> ExcelReader addModel(Class<T> modelCla, int sheetIndex) {
		return addModel(ExcelReaderModelParam.<T>builder()
				.modelCla(modelCla)
				.sheetIndex(sheetIndex)
				.build());
	}

	/**
	 * 添加list 解析
	 * @param param
	 * @return
	 */
	<T extends ExcelBaseModel> ExcelReader addList(ExcelReaderListParam<T> param);

	default <T extends ExcelBaseModel> ExcelReader addList(Class<T> modelCla, int sheetIndex) {
		return addList(ExcelReaderListParam.<T>builder()
				.modelCla(modelCla)
				.sheetIndex(sheetIndex)
				.build());
	}

	/**
	 * 获取单个bean
	 *
	 * @param modelCla
	 * @return
	 */
	<T extends ExcelBaseModel> T getModel(Class<T> modelCla);

	/**
	 * 获取bean list
	 *
	 * @param modelCla
	 * @return
	 */
	<T extends ExcelBaseModel> List<T> getList(Class<T> modelCla);


	/**
	 * 执行解析，遇到错误中断
	 */
	ExcelReader parse();

	/**
	 * 自定义读
	 * @param customReader
	 * @return
	 */
	ExcelReader readCustom(ExcelCustomReader customReader);

	/**
	 * 解析并获取模型
	 * @param modelCla
	 * @return
	 */
	<T extends ExcelBaseModel> T parseAndGetModel(Class<T> modelCla);

	/**
	 * 解析并获取模型list
	 * @param modelCla
	 * @return
	 */
	<T extends ExcelBaseModel> List<T> parseAndGetList(Class<T> modelCla);
	/**
	 * 设置是否快速失败
	 * @param failFast 是否快速失败,true 快速失败将抛出异常，false不快速失败，将继续执行直到最后通过
	 */
	ExcelReader setFailFast(Boolean failFast);

	/**
	 * 获取错误消息
	 * @return 错误消息，如果没有错误消息返回空字符串不是null
	 */
	String getErrorMsg();
}
