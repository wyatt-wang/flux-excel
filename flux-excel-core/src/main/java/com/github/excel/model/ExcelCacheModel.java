package com.github.excel.model;

import com.github.excel.annotation.ExcelWrite;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出缓存model
 */
@Data
public class ExcelCacheModel {
	/**
	 * excelExport 注解
	 */
	private ExcelWrite excelWrite;
	/**
	 * 字段列表
	 */
	private List<ExcelCacheFieldModel> fieldModelList;
	/**
	 * 字段map
	 */
	private Map<String, ExcelCacheFieldModel> fieldModelMap;

}
