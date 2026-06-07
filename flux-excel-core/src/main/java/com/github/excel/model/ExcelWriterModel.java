package com.github.excel.model;

import lombok.Data;

import java.util.List;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: 导出model
 */
@Data
public class ExcelWriterModel {
	private String sheetName;
	private ExcelBaseModel dataModel;
	private List<? extends ExcelBaseModel> dataModelList;
	private ExcelCacheModel cacheModel;
	private Class<? extends ExcelBaseModel> excelModelClass;
	private Integer rowIndex;
	private Integer colIndex;
	private Boolean fillTemplate;
}
