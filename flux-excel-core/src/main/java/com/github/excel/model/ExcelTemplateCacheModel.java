package com.github.excel.model;

import com.github.excel.boot.WorkbookCachePool;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @Author: Vachel Wang
 * @Date: 2026/4/24
 * @Description: Excel 文件缓存model
 */
@AllArgsConstructor
public class ExcelTemplateCacheModel {

	@Getter
	private final WorkbookCachePool.WorkbookCacheModel workbookThreadLocal;

}
