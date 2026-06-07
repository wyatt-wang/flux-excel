package com.github.excel.model;

import com.github.excel.exception.ExcelReaderException;
import lombok.Data;

import jakarta.validation.groups.Default;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 模型通用基础类
 */
@Data
public class ExcelBaseModel{

	private Map<String,String> modelColAddress = new ConcurrentHashMap<>();

	/**
	 * 执行检查方法由子类重写
	 * @throws ExcelReaderException
	 */
	public void callback() throws ExcelReaderException {

	}

	/**
	 * 执行检查方法由子类重写
	 * @throws ExcelReaderException
	 */
	public Class<?>[] validationGroup() {
		return new Class[]{Default.class};
	}
}
