package com.github.excel.write;

import com.github.excel.model.ExcelBaseModel;
import com.github.excel.write.style.AbstractExcelStyle;

import java.io.File;
import java.io.OutputStream;
import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出大数据list
 */
public interface ExcelLargeListWriter {

	/**
	 * 添加modelList，list对象类型相同
	 * 请勿传递不同类型的对象
	 * @param modelList
	 * @param <T>
	 */
	<T extends ExcelBaseModel> void  process(List<T> modelList, String[] excludeFields, Class<? extends ExcelBaseModel> modelCla);

	/**
	 * 导出list
	 * @param modelList
	 * @param <T>
	 */
	<T extends ExcelBaseModel> void  process(List<T> modelList, Class<? extends ExcelBaseModel> modelCla);

	/**
	 * 设置没有数据提示
	 * @param noneDataTips
	 */
	void setNoneDataTips(boolean noneDataTips);

	/**
	 * 执行导出，导出到文件
	 * @param outputStream
	 */
	void export(OutputStream outputStream);

	/**
	 * 执行导出，导出到客户端
	 * @param file
	 * @param fileName
	 */
	void export(File file, String fileName);

	/**
	 * 添加样式
	 * @param styleClass
	 */
	void addStyle(Class<? extends AbstractExcelStyle> styleClass);
	/**
	 * 设置sheet名称
	 * @param sheetName sheet名称
	 * @return
	 */
	void setSheetName(String sheetName);

	void setCla(Class<? extends ExcelBaseModel> cla);

	/**
	 * 关闭
	 */
	void close() ;


}
