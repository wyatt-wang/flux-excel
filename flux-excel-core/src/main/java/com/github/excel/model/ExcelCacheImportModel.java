package com.github.excel.model;

import com.github.excel.annotation.ExcelValidation;
import com.github.excel.annotation.ExcelRead;
import com.github.excel.annotation.ExcelReadProperty;
import lombok.Data;

import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导入缓存model
 */
@Data
public class ExcelCacheImportModel {
	private ExcelRead excelRead;
	private ExcelValidation validation;
	private Map<String, ExcelCacheImportFieldModel> fieldModelMap;
	private int maxHeadDepth = 1;

	@Data
	public static class ExcelCacheImportFieldModel {
		private String field;
		private ExcelReadProperty importProperty;
		private Method setMethod;
		private List<String> headNames;
	}
}
